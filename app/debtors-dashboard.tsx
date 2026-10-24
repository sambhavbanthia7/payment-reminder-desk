"use client"

import { useCallback, useEffect, useMemo, useRef, useState } from "react"
import Image from "next/image"
import { ArrowDownLeft, ArrowRight, ArrowUpRight, Bell, BookOpen, CalendarDays, CircleAlert, Clock3, Cloud, Download, FileCheck2, FileSpreadsheet, Flag, History, LayoutDashboard, Mail, Search, Send, Settings2, ShieldAlert, Upload, Users, Wallet, X } from "lucide-react"
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle } from "@/components/ui/dialog"
import { Toaster } from "@/components/ui/sonner"
import { toast } from "sonner"
import { addActivity, ageingLabels, allocationKey, calculateInvoices, dayNumber, displayDate, groupParties, money, partyKey, recordReceipt, reminderQueue, renderMessage, shortMoney, today, updateContact, validDate, validEmail, type Confirmation, type Invoice, type Party, type Receipt, type Reminder, type Settings, type WorkbookState } from "@/lib/debtors"
import { exportDebtors, importDebtors } from "@/lib/debtors-workbook"
import { applySendResults, staleReason, type EmailStatus, type SendResult } from "@/lib/email-dispatch"
import { readDraft, readLinkedFile, writeDraft, writeLinkedFile, type LinkedFile, type RecoveryDraft } from "@/lib/debtors-storage"
import { chooseDriveFile, createDriveTracker, disconnectDrive, downloadTracker, DriveConflict, DriveLocked, fetchDriveStatus, listDriveFiles, openDecision, pollDecision, saveTracker, type DriveBase, type DriveFile, type OneDriveStatus } from "@/lib/onedrive-sync"
import { changedOutside, fileDecision, writeExcelFile, type ExcelHandle } from "@/lib/excel-file"

type Page = "overview" | "queue" | "balances" | "invoices" | "ageing" | "risk" | "logs" | "quality"
type FileHandle = ExcelHandle
type Modal = { kind:"invoice"; id:string } | { kind:"contact"; party:string } | { kind:"payment"; invoice:string; receipt?:string } | { kind:"followup"; id:string } | { kind:"confirmation"; party:string } | { kind:"message"; party:Party; type:Reminder["type"] } | { kind:"draft"; id:string } | { kind:"settings" } | { kind:"onedrive" } | null
const titles:Record<Page,string>={overview:"Receivables overview",queue:"Today’s reminder queue",balances:"Balance confirmation",invoices:"Invoice register",ageing:"Ageing analysis",risk:"Critical debtors",logs:"Activity & reminder history",quality:"Data checks"}
const navigation:[Page,string,typeof LayoutDashboard][]=[["overview","Overview",LayoutDashboard],["queue","Reminder queue",Bell],["balances","Balance confirmation",FileCheck2],["invoices","Invoices & collections",BookOpen],["ageing","Ageing analysis",Clock3],["risk","Risk & escalation",ShieldAlert],["logs","History & logs",History],["quality","Data checks",CircleAlert]]

export default function DebtorsDashboard({account}:{account:string}) {
  const [state,setState]=useState<WorkbookState|null>(null)
  const [source,setSource]=useState<ArrayBuffer|null>(null)
  const [restoring,setRestoring]=useState(true)
  const [busy,setBusy]=useState(false)
  const [page,setPage]=useState<Page>("overview")
  const [date,setDate]=useState(today)
  const [search,setSearch]=useState("")
  const [area,setArea]=useState("All areas")
  const [rse,setRse]=useState("All RSEs")
  const [selected,setSelected]=useState<Set<string>>(new Set())
  const [modal,setModal]=useState<Modal>(null)
  const [exported,setExported]=useState(0)
  const [autoSave,setAutoSave]=useState(false)
  const [autoRevision,setAutoRevision]=useState(-1)
  const [excelStatus,setExcelStatus]=useState("")
  const [draftStatus,setDraftStatus]=useState("Recovery draft not saved yet")
  const [invoiceFilter,setInvoiceFilter]=useState("All invoices")
  const [logTab,setLogTab]=useState("Activity")
  const [mail,setMail]=useState<EmailStatus>({configured:false})
  const [sending,setSending]=useState(false)
  // Everyone signs in with the same mailbox account, so each computer remembers who is sending for the log.
  const [senderName,setSenderName]=useState(()=>{try{return typeof window==="undefined"?"":localStorage.getItem("jal-sender-name")??""}catch{return ""}})
  const askSenderName=()=>{
    const value=window.prompt("Your name. It is recorded in the log with each email you send from this computer.",senderName)?.replace(/\s+/g," ").trim()
    if(value===undefined)return null
    if(!value){toast.error("Enter your name so the log shows who sent each email");return null}
    try{localStorage.setItem("jal-sender-name",value)}catch{}
    setSenderName(value);return value
  }
  const fileInput=useRef<HTMLInputElement>(null)
  const fileHandle=useRef<FileHandle|null>(null)
  const fileStamp=useRef<{size:number;modified:number}|null>(null)
  const persistence=useRef<Promise<unknown>>(Promise.resolve())
  const fileWrites=useRef<Promise<unknown>>(Promise.resolve())
  // OneDrive: `driveBase` is set while the open workbook is the shared OneDrive tracker.
  const [drive,setDrive]=useState<OneDriveStatus>({configured:false,connected:false})
  const [driveBase,setDriveBase]=useState<DriveBase|null>(null)
  const [driveSync,setDriveSync]=useState<{state:"idle"|"saving"|"saved"|"conflict"|"locked"|"error";message?:string;at?:string;remote?:DriveFile;failedRevision?:number}>({state:"idle"})
  const driveWrites=useRef<Promise<unknown>>(Promise.resolve())
  // Latest values for async OneDrive work that outlives a render.
  const latest=useRef({state,source,driveBase,date,exported})
  useEffect(()=>{latest.current={state,source,driveBase,date,exported}})
  // Excel auto-save: the tracker file this computer saves into, remembered between visits.
  const [linkedName,setLinkedName]=useState("")
  const [pendingFile,setPendingFile]=useState<LinkedFile|null>(null)
  const [fileConflict,setFileConflict]=useState(false)
  const [savedAt,setSavedAt]=useState("")
  const [retryTick,setRetryTick]=useState(0)

  const restoreDraft=(draft:RecoveryDraft)=>{
    const recovered=draft.state
    if(!recovered.activity.length){recovered.activity=[addActivity(recovered,"Workbook opened",null,`${recovered.invoices.length} invoices · ${recovered.receipts.length} receipts`)];recovered.revision+=1}
    setState(recovered);setSource(draft.bytes);setExported(draft.exportedRevision);setDraftStatus("Recovered on this browser")
  }
  /** Load the shared tracker from OneDrive. Opening it is not logged, so viewing never creates a save. */
  const openFromDrive=useCallback(async(reset:boolean)=>{
    const {bytes,file}=await downloadTracker()
    const next=importDebtors(new Uint8Array(bytes),file.name)
    next.revision=1
    const base:DriveBase={itemId:file.id,eTag:file.eTag,cTag:file.cTag,savedRevision:1}
    await persistence.current.catch(()=>{})
    await writeDraft({bytes,state:next,exportedRevision:1,savedAt:new Date().toISOString(),drive:base}).catch(()=>{})
    latest.current.driveBase=base
    setSource(bytes);setState(next);setExported(1);setAutoSave(false);setAutoRevision(-1);fileHandle.current=null;fileStamp.current=null
    setDriveBase(base);setDriveSync({state:"saved",at:file.modified});setDraftStatus("Opened from OneDrive")
    if(reset){setArea("All areas");setRse("All RSEs");setSearch("");setSelected(new Set());setPage("overview");setModal(null)}
    return file
  },[])
  /** Open the tracker straight from a file on this computer and save every change back into it. */
  const openFromHandle=useCallback(async(handle:FileHandle,reset:boolean)=>{
    const file=await handle.getFile()
    if(file.size>20*1024*1024)throw new Error("The tracker must be smaller than 20 MB")
    const bytes=await file.arrayBuffer(),next=importDebtors(new Uint8Array(bytes),file.name)
    // Opening is not logged, so simply opening the file never rewrites it.
    next.revision=1
    const stamp={size:file.size,modified:file.lastModified}
    await persistence.current.catch(()=>{})
    await writeDraft({bytes,state:next,exportedRevision:1,savedAt:new Date().toISOString()}).catch(()=>{})
    await writeLinkedFile({handle,name:file.name,stamp}).catch(()=>{})
    fileHandle.current=handle;fileStamp.current=stamp
    setSource(bytes);setState(next);setExported(1);setAutoSave(true);setAutoRevision(1);setLinkedName(file.name);setFileConflict(false);setPendingFile(null)
    setSavedAt(new Date(file.lastModified).toISOString());setExcelStatus("");setDriveBase(null);setDriveSync({state:"idle"});setDraftStatus(`Opened from ${file.name}`)
    if(reset){setArea("All areas");setRse("All RSEs");setSearch("");setSelected(new Set());setPage("overview");setModal(null)}
    return file.name
  },[])
  /** Pick up the remembered tracker file again, taking in any edits made to it outside the app. */
  const reconnectFile=useCallback(async(linked:LinkedFile,local:{state:WorkbookState|null;exported:number})=>{
    const file=await linked.handle.getFile()
    const current={size:file.size,modified:file.lastModified}
    const unsaved=Boolean(local.state&&local.state.revision>local.exported)
    if(!local.state||local.state.fileName!==linked.name){
      if(unsaved&&!window.confirm(`Open ${linked.name}? This computer has unsaved changes to another workbook, and they will be replaced.`))return
      await openFromHandle(linked.handle,false);return
    }
    const decision=fileDecision(linked.stamp,current,unsaved)
    if(decision==="reload"){await openFromHandle(linked.handle,false);toast.info(`Loaded the latest ${linked.name}. It was changed outside the app.`);return}
    fileHandle.current=linked.handle;fileStamp.current=decision==="same"?current:linked.stamp
    setLinkedName(linked.name);setPendingFile(null);setAutoRevision(local.exported);setSavedAt(new Date(file.lastModified).toISOString())
    if(decision==="conflict"){setFileConflict(true);setAutoSave(false)}else{setFileConflict(false);setAutoSave(true)}
  },[openFromHandle])
  useEffect(()=>{
    let cancelled=false
    ;(async()=>{
      let draft:RecoveryDraft|null=null
      try{draft=await readDraft()}catch{setDraftStatus("Browser recovery is unavailable. Save the workbook after changes.")}
      if(draft&&draft.state.version!==1)draft=null
      const status=await fetchDriveStatus().catch(()=>({configured:false,connected:false} as OneDriveStatus))
      if(cancelled)return
      setDrive(status)
      const file=status.file
      if(!file){
        if(draft)restoreDraft(draft)
        const linked=await readLinkedFile().catch(()=>null)
        if(!linked||cancelled)return
        // Chrome may keep the permission between visits; otherwise one click on "Continue" is needed.
        const permission=await linked.handle.queryPermission?.({mode:"readwrite"}).catch(()=>"prompt")
        if(permission==="granted"){try{await reconnectFile(linked,{state:draft?.state??null,exported:draft?.exportedRevision??0})}catch{setPendingFile(linked)}}
        else setPendingFile(linked)
        return
      }
      const decision=openDecision(draft?{revision:draft.state.revision,exportedRevision:draft.exportedRevision,drive:draft.drive}:null,file)
      if(decision==="use-local"&&draft){restoreDraft(draft);setDriveBase({...draft.drive!,eTag:file.eTag});setDriveSync({state:"saved",at:file.modified});return}
      if(decision==="ask"&&draft&&!window.confirm(`The shared tracker in OneDrive has newer changes${file.modifiedBy?` by ${file.modifiedBy}`:""}, and this computer also has changes that were not saved to OneDrive.\n\nOK: open the OneDrive version. This computer's unsaved changes are discarded.\nCancel: keep this computer's version for now. You can then download it or replace the OneDrive file.`)){
        restoreDraft(draft)
        if(draft.drive?.itemId===file.id){setDriveBase(draft.drive);setDriveSync({state:"conflict",remote:file})}
        return
      }
      try{await openFromDrive(false)}catch(error){if(draft)restoreDraft(draft);toast.error(error instanceof Error?error.message:"The OneDrive tracker could not be opened")}
    })().finally(()=>{if(!cancelled)setRestoring(false)})
    return()=>{cancelled=true}
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[])
  useEffect(()=>{
    if(!state||!source)return
    let active=true
    const timer=setTimeout(()=>{
      persistence.current=persistence.current.catch(()=>{}).then(()=>writeDraft({state,bytes:source,exportedRevision:exported,savedAt:new Date().toISOString(),drive:driveBase??undefined})).then(()=>{if(active)setDraftStatus("Recovery draft saved on this browser")}).catch(()=>{if(active)setDraftStatus("Recovery failed. Save your workbook now.")})
    },350)
    return()=>{active=false;clearTimeout(timer)}
  },[state,source,exported,driveBase])
  useEffect(()=>{
    const warn=(event:BeforeUnloadEvent)=>{if(state&&state.revision>(driveBase?driveBase.savedRevision:autoSave?autoRevision:exported)){event.preventDefault();event.returnValue=""}}
    window.addEventListener("beforeunload",warn);return()=>window.removeEventListener("beforeunload",warn)
  },[state,exported,autoSave,autoRevision,driveBase])
  useEffect(()=>{
    let cancelled=false
    fetch("/api/email/status").then((r)=>r.json() as Promise<EmailStatus>).then((status)=>{if(!cancelled)setMail(status)}).catch(()=>{if(!cancelled)setMail({configured:false,problem:"Email status unavailable"})})
    return()=>{cancelled=true}
  },[])
  const all=useMemo(()=>state?calculateInvoices(state,date):[],[state,date])
  const filtered=useMemo(()=>all.filter((i)=>(area==="All areas"||i.area===area)&&(rse==="All RSEs"||i.rse===rse)&&(!search||[i.party,i.number,i.area,i.rse,i.email].some((v)=>v.toLowerCase().includes(search.toLowerCase())))),[all,area,rse,search])
  const parties=useMemo(()=>groupParties(filtered.filter((i)=>i.outstanding>0)),[filtered])
  const balanceParties=useMemo(()=>{const keys=new Set(filtered.map((i)=>partyKey(i.party)));return groupParties(all.filter((i)=>i.outstanding>0)).filter((p)=>keys.has(p.key))},[all,filtered])
  const queue=useMemo(()=>state?reminderQueue(state,filtered,date):[],[state,filtered,date])
  const invoiceIds=new Set(all.map((i)=>allocationKey(i.party,i.number)))
  const unmatched=state?.receipts.filter((r)=>!invoiceIds.has(allocationKey(r.party,r.invoice))) ?? []
  const open=filtered.filter((i)=>i.outstanding>0)
  const total=open.reduce((s,i)=>s+i.outstanding,0)
  const overdue=open.filter((i)=>i.overdue>0)
  const critical=parties.filter((p)=>p.risk==="Critical")
  const issues=all.filter((i)=>i.issues.length>0)
  const areas=[...new Set(all.map((i)=>i.area).filter(Boolean))].sort()
  const rses=[...new Set(all.map((i)=>i.rse).filter(Boolean))].sort()
  const autoSaveAvailable=typeof window!=="undefined"&&typeof (window as unknown as {showSaveFilePicker?:unknown}).showSaveFilePicker==="function"
  const dirty=Boolean(state&&state.revision>(driveBase?driveBase.savedRevision:autoSave?autoRevision:exported))
  const go=(next:Page)=>{setPage(next);setSelected(new Set());setInvoiceFilter("All invoices")}
  const commit=(next:WorkbookState)=>setState(next)
  const loadFile=async(file:File)=>{
    if(dirty&&!window.confirm("There are changes not exported to Excel. Replace this working workbook? Cancel to save it first."))return
    setBusy(true)
    try {
      if(file.size>20*1024*1024)throw new Error("Upload a workbook smaller than 20 MB")
      const bytes=await file.arrayBuffer(),next=importDebtors(new Uint8Array(bytes),file.name)
      next.activity.push(addActivity(next,"Workbook opened",null,`${next.invoices.length} invoices · ${next.receipts.length} receipts`));next.revision=1
      await persistence.current.catch(()=>{})
      let recoverySaved=true
      await writeDraft({bytes,state:next,exportedRevision:0,savedAt:new Date().toISOString()}).catch(()=>{recoverySaved=false})
      setSource(bytes);setState(next);setExported(0);setAutoSave(false);setAutoRevision(-1);setExcelStatus("");setArea("All areas");setRse("All RSEs");setSearch("");setSelected(new Set());setPage("overview");setModal(null);setDraftStatus(recoverySaved?"Recovery draft saved on this browser":"Browser recovery is unavailable. Save the workbook after changes.");fileHandle.current=null;fileStamp.current=null;setDriveBase(null);setDriveSync({state:"idle"});setLinkedName("");setPendingFile(null);setFileConflict(false)
      toast.success(`${next.invoices.length.toLocaleString("en-IN")} invoices loaded from ${file.name}`)
    }catch(error){toast.error(error instanceof Error?error.message:"This workbook could not be loaded")}finally{setBusy(false);if(fileInput.current)fileInput.current.value=""}
  }
  const baseName=state?.fileName.replace(/(?:-updated(?:-\d{4}-\d{2}-\d{2})?)+\.xlsx$/i,".xlsx") ?? "Debtors-Tracker.xlsx"
  const writeConnectedFile=useCallback((handle:FileHandle,bytes:ArrayBuffer,snapshot:WorkbookState,reviewDate:string)=>{
    const operation=fileWrites.current.catch(()=>{}).then(async()=>{
      if(fileHandle.current!==handle)return false
      const output=exportDebtors(new Uint8Array(bytes),snapshot,reviewDate)
      const stamp=await writeExcelFile(handle,output,fileStamp.current)
      if(fileHandle.current!==handle)return false
      fileStamp.current=stamp
      const name=(await handle.getFile()).name
      void writeLinkedFile({handle,name,stamp}).catch(()=>{})
      setLinkedName(name);setSavedAt(new Date().toISOString())
      return true
    });fileWrites.current=operation;return operation
  },[])
  const save=async(direct:boolean)=>{
    if(!state||!source)return
    setBusy(true)
    try{
      let handle=fileHandle.current
      const picker=(window as unknown as {showSaveFilePicker?:(options:unknown)=>Promise<FileHandle>}).showSaveFilePicker
      if(direct&&!handle){
        if(!picker)throw new Error("Direct file saving is unavailable in this browser. Use Download workbook instead.")
        handle=await picker({suggestedName:baseName.replace(/\.xlsx$/i,"-updated.xlsx"),types:[{description:"Excel workbook",accept:{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":[".xlsx"]}}]})
      }
      if(direct&&handle){fileHandle.current=handle;fileStamp.current=null;if(await writeConnectedFile(handle,source,state,date)){setAutoSave(true);setAutoRevision(state.revision);setExcelStatus("Excel file saved. Future changes will be written automatically.");setExported((r)=>Math.max(r,state.revision));toast.success("Excel auto-save enabled, including Activity Log and Reminder Log")}}
      else{const bytes=exportDebtors(new Uint8Array(source),state,date),blob=new Blob([bytes.slice().buffer as ArrayBuffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),url=URL.createObjectURL(blob),link=document.createElement("a");link.href=url;link.download=baseName.replace(/\.xlsx$/i,`-updated-${today()}.xlsx`);link.click();setTimeout(()=>URL.revokeObjectURL(url),30000);setExported((r)=>Math.max(r,state.revision));setExcelStatus("Downloaded Excel includes Activity Log and Reminder Log. Open the downloaded copy to see them.");toast.success("Excel downloaded with log sheets. It opens on Activity Log.")}
    }catch(error){if(!(error instanceof Error&&error.name==="AbortError")){setAutoSave(false);setExcelStatus("Excel was not saved. Download a copy to keep your latest logs.");toast.error(error instanceof Error?error.message:"Workbook could not be saved")}}finally{setBusy(false)}
  }
  useEffect(()=>{
    const handle=fileHandle.current
    if(!autoSave||!handle||!state||!source||state.revision<=autoRevision)return
    let cancelled=false
    const timer=setTimeout(()=>{void writeConnectedFile(handle,source,state,date).then((written)=>{if(written){setAutoRevision((r)=>Math.max(r,state.revision));setExported((r)=>Math.max(r,state.revision));setExcelStatus("")}}).catch((error:unknown)=>{
      if(cancelled&&fileHandle.current!==handle)return
      if(changedOutside(error)){setFileConflict(true);setAutoSave(false);setExcelStatus("");toast.error("The tracker was changed outside the app. Choose which version to keep.")}
      // Usually the file is open in Excel, which blocks writing. Keep the changes and try again shortly.
      else{setExcelStatus("Could not save into the Excel file. If it is open in Excel, close it there. Trying again in 15 seconds.");setTimeout(()=>setRetryTick((t)=>t+1),15000)}
    })},800)
    return()=>{cancelled=true;clearTimeout(timer)}
  },[state,source,date,autoSave,autoRevision,writeConnectedFile,retryTick])
  useEffect(()=>{
    // Notice edits made to the file in Excel while the app is open.
    if(!autoSave)return
    let checking=false
    const check=async()=>{
      const handle=fileHandle.current
      if(!handle||checking||document.hidden)return
      checking=true
      try{
        await fileWrites.current.catch(()=>{})
        const file=await handle.getFile().catch(()=>null)
        if(!file||fileHandle.current!==handle)return
        const {state:current,exported:saved}=latest.current
        const decision=fileDecision(fileStamp.current,{size:file.size,modified:file.lastModified},Boolean(current&&current.revision>saved))
        if(decision==="reload"){await openFromHandle(handle,false);toast.info(`Loaded the latest ${file.name}. It was changed outside the app.`)}
        else if(decision==="conflict"){setFileConflict(true);setAutoSave(false)}
      }finally{checking=false}
    }
    const timer=setInterval(()=>void check(),30000)
    const onFocus=()=>void check()
    window.addEventListener("focus",onFocus)
    return()=>{clearInterval(timer);window.removeEventListener("focus",onFocus)}
  },[autoSave,openFromHandle])
  /** Save the open workbook over the OneDrive tracker. Writes run one at a time. */
  const pushToDrive=useCallback((force:boolean)=>{
    const operation=driveWrites.current.catch(()=>{}).then(async()=>{
      const {state:snapshot,source:bytes,driveBase:base,date:reviewDate}=latest.current
      if(!snapshot||!bytes||!base||(!force&&snapshot.revision<=base.savedRevision))return
      setDriveSync({state:"saving"})
      try{
        const file=await saveTracker(exportDebtors(new Uint8Array(bytes),snapshot,reviewDate),base,force)
        const next:DriveBase={itemId:file.id,eTag:file.eTag,cTag:file.cTag,savedRevision:snapshot.revision}
        latest.current.driveBase=next;setDriveBase(next);setDriveSync({state:"saved",at:file.modified});setExported((r)=>Math.max(r,snapshot.revision))
        if(force)toast.success("OneDrive now has this computer's version")
      }catch(error){
        if(error instanceof DriveConflict)setDriveSync({state:"conflict",remote:error.file})
        else if(error instanceof DriveLocked){setDriveSync({state:"locked",message:error.message,failedRevision:snapshot.revision});setTimeout(()=>{setDriveSync((s)=>s.state==="locked"?{state:"idle"}:s)},15000)}
        else setDriveSync({state:"error",message:error instanceof Error?error.message:"OneDrive save failed",failedRevision:snapshot.revision})
      }
    });driveWrites.current=operation;return operation
  },[])
  useEffect(()=>{
    if(!driveBase||!state||!source||state.revision<=driveBase.savedRevision)return
    // After a conflict nothing is saved until someone chooses which version to keep. After a failure, the next change retries.
    if(driveSync.state==="conflict"||driveSync.state==="saving"||((driveSync.state==="error"||driveSync.state==="locked")&&driveSync.failedRevision===state.revision))return
    const timer=setTimeout(()=>void pushToDrive(false),1500)
    return()=>clearTimeout(timer)
  },[state,source,driveBase,driveSync,pushToDrive])
  useEffect(()=>{
    if(!driveBase?.itemId)return
    let checking=false
    const check=async()=>{
      if(checking||document.hidden)return
      checking=true
      try{
        await driveWrites.current.catch(()=>{})
        const status=await fetchDriveStatus().catch(()=>null)
        if(!status)return
        setDrive(status)
        const base=latest.current.driveBase,current=latest.current.state
        if(!status.file||!base||!current)return
        const decision=pollDecision(base,current.revision,status.file)
        if(decision==="same"){if(status.file.eTag!==base.eTag){const next={...base,eTag:status.file.eTag};latest.current.driveBase=next;setDriveBase(next)}return}
        if(decision==="conflict"){setDriveSync({state:"conflict",remote:status.file});return}
        await openFromDrive(false)
        toast.info(`Updated with the latest tracker from OneDrive${status.file.modifiedBy?` (last changed by ${status.file.modifiedBy})`:""}`)
      }finally{checking=false}
    }
    const timer=setInterval(()=>void check(),30000)
    const onFocus=()=>void check()
    window.addEventListener("focus",onFocus)
    return()=>{clearInterval(timer);window.removeEventListener("focus",onFocus)}
  },[driveBase?.itemId,openFromDrive])
  useEffect(()=>{
    // The OneDrive sign-in returns here with ?onedrive=connected or ?onedrive=failed.
    if(restoring)return
    const result=new URLSearchParams(window.location.search).get("onedrive")
    if(!result)return
    window.history.replaceState(null,"",window.location.pathname)
    setTimeout(()=>{
      if(result==="connected"){toast.success("OneDrive connected");if(!drive.file)setModal({kind:"onedrive"})}
      else toast.error("OneDrive sign-in did not complete. Try Connect OneDrive again.")
    },0)
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[restoring])
  const resolveConflict=async(choice:"theirs"|"mine")=>{
    if(choice==="theirs"){
      if(!window.confirm("Open the OneDrive version? Changes made on this computer since the last save to OneDrive will be discarded. Download a copy first if you need them."))return
      try{await openFromDrive(false);toast.success("Opened the latest OneDrive version")}catch(error){toast.error(error instanceof Error?error.message:"OneDrive could not be reached")}
    }else{
      if(!window.confirm(`Replace the OneDrive tracker with this computer's version? Changes saved in OneDrive${driveSync.remote?.modifiedBy?` by ${driveSync.remote.modifiedBy}`:""} since you opened it will be lost.`))return
      await pushToDrive(true)
    }
  }
  const useDriveFile=async(file:DriveFile)=>{
    if(dirty&&!window.confirm("This computer has changes that are not saved. Open the OneDrive tracker anyway? Download a copy first if you need them."))return
    setBusy(true)
    try{await chooseDriveFile(file.id);setDrive(await fetchDriveStatus());await openFromDrive(true);toast.success(`Now using ${file.name} from OneDrive`)}
    catch(error){toast.error(error instanceof Error?error.message:"That tracker could not be opened")}finally{setBusy(false)}
  }
  const putIntoDrive=async()=>{
    if(!state||!source)return
    setBusy(true)
    try{
      const file=await createDriveTracker(exportDebtors(new Uint8Array(source),state,date),baseName)
      const base:DriveBase={itemId:file.id,eTag:file.eTag,cTag:file.cTag,savedRevision:state.revision}
      latest.current.driveBase=base;setDriveBase(base);setDriveSync({state:"saved",at:file.modified});setExported((r)=>Math.max(r,state.revision));setAutoSave(false);fileHandle.current=null
      setDrive(await fetchDriveStatus());setModal(null);toast.success(`${file.name} is now the shared tracker in OneDrive`)
    }catch(error){toast.error(error instanceof Error?error.message:"The workbook could not be put into OneDrive")}finally{setBusy(false)}
  }
  const disconnect=async()=>{
    if(!window.confirm("Disconnect OneDrive for everyone? The tracker stays in OneDrive, but the app stops opening and saving it until someone connects again."))return
    await disconnectDrive();setDrive(await fetchDriveStatus().catch(()=>({configured:true,connected:false})));setDriveBase(null);setDriveSync({state:"idle"});setModal(null);toast.success("OneDrive disconnected")
  }
  /** Open the tracker from this computer. Where the browser allows, the app then saves straight into it. */
  const openTrackerFile=async()=>{
    const openPicker=(window as unknown as {showOpenFilePicker?:(options:unknown)=>Promise<FileHandle[]>}).showOpenFilePicker
    if(!openPicker){fileInput.current?.click();return}
    if(dirty&&!window.confirm("There are changes not saved to Excel. Open a different tracker anyway?"))return
    let handle:FileHandle
    try{[handle]=await openPicker.call(window,{multiple:false,types:[{description:"Excel workbook",accept:{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":[".xlsx"]}}]})}
    catch(error){if(!(error instanceof Error&&error.name==="AbortError"))toast.error("The file could not be opened");return}
    setBusy(true)
    try{
      // Ask for permission to save while this click still counts, so no second prompt is needed.
      const granted=!handle.requestPermission||await handle.requestPermission({mode:"readwrite"})==="granted"
      const name=await openFromHandle(handle,true)
      if(granted)toast.success(`${name} is open. Changes are saved into this file automatically.`)
      else{setAutoSave(false);toast.error("Saving into the file was not allowed. Use Download Excel + logs to keep changes.")}
    }catch(error){toast.error(error instanceof Error?error.message:"This workbook could not be loaded")}finally{setBusy(false)}
  }
  const continueWithFile=async()=>{
    if(!pendingFile)return
    try{
      const permission=await pendingFile.handle.requestPermission?.({mode:"readwrite"})
      if(permission&&permission!=="granted")return void toast.error("The app needs permission to save into the file")
      await reconnectFile(pendingFile,{state,exported})
    }catch{
      toast.error(`${pendingFile.name} could not be opened. It may have been moved or renamed. Open it again.`)
      setPendingFile(null);void writeLinkedFile(null).catch(()=>{})
    }
  }
  const resolveFileConflict=async(choice:"file"|"mine")=>{
    const handle=fileHandle.current
    if(!handle)return
    if(choice==="file"){
      if(!window.confirm("Open the version in the Excel file? Changes made in the app since the last save will be discarded. Download a copy first if you need them."))return
      try{await openFromHandle(handle,false);toast.success("Opened the latest version of the file")}catch(error){toast.error(error instanceof Error?error.message:"The file could not be opened")}
    }else{
      if(!window.confirm("Save the app's version into the Excel file? Edits made to the file outside the app since it was opened will be lost."))return
      fileStamp.current=null;setFileConflict(false);setAutoRevision((r)=>Math.min(r,(state?.revision??1)-1));setAutoSave(true)
    }
  }
  const signOut=async()=>{
    if(dirty&&!window.confirm("Some changes are not saved yet. Sign out anyway? They stay in this browser's recovery copy."))return
    await fetch("/api/auth/logout",{method:"POST"}).catch(()=>{})
    window.location.reload()
  }
  const toggle=(key:string)=>setSelected((current)=>{const next=new Set(current);if(next.has(key))next.delete(key);else next.add(key);return next})
  const prepare=(rows:Party[],type:Reminder["type"],single=false)=>{
    if(!state)return
    const allowed=rows.filter((p)=>(single||selected.has(p.key))&&p.email&&!p.issues.includes("Duplicate party + invoice")&&!state.reminders.some((r)=>r.type===type&&r.date===date&&r.status!=="Cancelled"&&r.status!=="Failed"&&r.invoiceIds.some((id)=>p.invoices.some((i)=>i.id===id))))
    if(!allowed.length)return toast.error("Select at least one customer with a valid contact")
    if(!state.settings[type==="Reminder"?"reminderSubject":"balanceSubject"].trim()||!state.settings[type==="Reminder"?"reminderBody":"balanceBody"].trim())return toast.error("Add a subject and message in Email templates")
    const reminders:Reminder[]=allowed.map((p)=>({id:crypto.randomUUID(),at:new Date().toISOString(),date,type,party:p.name,email:p.email,invoiceIds:p.invoices.map((i)=>i.id),invoiceNumbers:p.invoices.map((i)=>i.number),amount:p.total,stage:[...p.invoices].sort((a,b)=>b.overdue-a.overdue)[0]?.stage ?? "",...renderMessage(state.settings,p,date,type),status:"Prepared"}))
    const events=reminders.map((r)=>({...addActivity(state,"Draft prepared",null,`${type} · ${r.invoiceNumbers.join(", ")} · ${money(r.amount)}`),party:r.party,after:r.id}))
    commit({...state,revision:state.revision+1,reminders:[...state.reminders,...reminders],activity:[...state.activity,...events]});setSelected(new Set());if(single)setModal(null);toast.success(`${reminders.length} drafts prepared. No email has been sent.`,{action:{label:"View drafts",onClick:()=>{setPage("logs");setLogTab("Reminder drafts")}}})
  }
  const cancelDraft=(reminder:Reminder)=>{if(!state)return;commit({...state,revision:state.revision+1,reminders:state.reminders.map((r)=>r.id===reminder.id?{...r,status:"Cancelled"}:r),activity:[...state.activity,{...addActivity(state,"Draft cancelled",null,reminder.subject),party:reminder.party,before:reminder.id}]});setModal(null);toast.success("Draft cancelled; eligible invoices can be queued again")}
  const sendDrafts=async(drafts:Reminder[])=>{
    if(!state||sending)return
    if(!mail.configured)return void toast.error(mail.problem||"Email sending is not set up yet")
    const name=senderName||askSenderName()
    if(!name)return
    const stale=drafts.map((d)=>({draft:d,reason:staleReason(d,all)})).filter((s)=>s.reason)
    const ready=drafts.filter((d)=>!stale.some((s)=>s.draft.id===d.id))
    if(stale.length)toast.warning(`${stale.length} draft${stale.length===1?"":"s"} not sent: ${stale[0].reason}. Cancel and prepare ${stale.length===1?"it":"them"} again.`)
    if(!ready.length)return
    const live=mail.mode==="live"
    if(!window.confirm(live?`Send ${ready.length} email${ready.length===1?"":"s"} to customers from ${mail.sender}? This cannot be undone.`:`TEST MODE: send ${ready.length} test email${ready.length===1?"":"s"} to ${mail.testRecipient}. Customers will not receive anything.`))return
    setSending(true)
    const results:SendResult[]=[]
    try{
      for(let start=0;start<ready.length;start+=10){
        const batch=ready.slice(start,start+10).map(({id,type,party,email,subject,body,invoiceNumbers,amount})=>({id,type,party,email,subject,body,invoiceNumbers,amount}))
        const response=await fetch("/api/email/send",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({drafts:batch,senderName:name})})
        const data=await response.json().catch(()=>({})) as {results?:SendResult[];skipped?:number;error?:string}
        if(response.status===401){toast.error(data.error??"Sign in again to send");setTimeout(()=>window.location.reload(),2500);return}
        if(!response.ok||!data.results)throw new Error(data.error??`Sending stopped (${response.status})`)
        results.push(...data.results)
        if(data.skipped)throw new Error(`The hourly sending limit was reached. ${ready.length-results.length} draft${ready.length-results.length===1?"":"s"} not sent yet; they stay Prepared`)
      }
    }catch(error){toast.error(`${error instanceof Error?error.message:"Sending stopped"}${results.length?` · ${results.length} already processed`:""}`)}
    finally{
      setSending(false)
      if(results.length){
        setState((current)=>current?applySendResults(current,results):current)
        const repeats=results.filter((r)=>r.repeat).length,failed=results.filter((r)=>r.status==="Failed"&&!r.repeat).length,accepted=results.length-failed-repeats
        const summary=[accepted?`${accepted} ${live?"sent":"test email"+(accepted===1?"":"s")+" sent to "+mail.testRecipient}`:"",repeats?`${repeats} already sent earlier, not sent again`:"",failed?`${failed} failed`:""].filter(Boolean).join(" · ")
        if(failed)toast.error(summary,{description:results.find((r)=>r.status==="Failed")?.detail});else toast.success(summary)
      }
    }
  }

  const input=<input ref={fileInput} type="file" accept=".xlsx" className="hidden" aria-label="Upload debtor tracker" onChange={(e)=>{if(e.target.files?.[0])void loadFile(e.target.files[0])}}/>
  if(restoring)return <div className="debtors loading-desk"><FileSpreadsheet size={30}/><p>Opening your workbook workspace…</p></div>
  if(!state)return <div className="debtors welcome"><Toaster position="top-right"/>{input}<div className="welcome-brand"><JalLogo/><span>Accounts receivable</span></div><div className="welcome-content"><span className="eyebrow">ACCOUNTS RECEIVABLE</span><h1>Every outstanding invoice.<br/>One clear next action.</h1><p>Open your Debtors Follow-up Tracker to review collections, prepare reminders, and keep the complete history in your Excel workbook.</p>{pendingFile&&<button className="desk-btn primary large" disabled={busy} onClick={()=>void continueWithFile()}><FileCheck2 size={18}/>Continue with {pendingFile.name}</button>}<button className={`desk-btn ${pendingFile?"":"primary "}large`} disabled={busy} onClick={()=>void openTrackerFile()}><Upload size={18}/>{busy?"Reading workbook…":pendingFile?"Open a different tracker":"Open debtor tracker"}</button>{drive.configured&&<button className="desk-btn large" disabled={busy} onClick={()=>setModal({kind:"onedrive"})}><Cloud size={18}/>{drive.connected?"Open from OneDrive":"Connect OneDrive"}</button>}<span className="welcome-hint">One .xlsx file · Debtors Data + Receipts · No column mapping</span><div className="welcome-features"><div><Clock3/><b>See what needs attention</b><p>Ageing, critical debtors and today’s follow-ups.</p></div><div><Wallet/><b>Keep payments reconciled</b><p>Record and correct receipts against each invoice.</p></div><div><History/><b>Carry your history forward</b><p>Save contacts, drafts and activity inside the workbook.</p></div></div><div className="welcome-note"><FileSpreadsheet size={20}/><p>Open the tracker once. In Chrome or Edge every change is then saved straight into that Excel file, and next time the app offers to continue with it.</p></div></div><Dialog open={modal?.kind==="onedrive"} onOpenChange={(open)=>{if(!open)setModal(null)}}><DialogContent className="debtors desk-dialog max-h-[92vh] overflow-y-auto sm:max-w-2xl" showCloseButton={false}><button className="dialog-close-action" aria-label="Close" onClick={()=>setModal(null)}><X size={20}/></button><DialogHeader><DialogTitle>OneDrive tracker</DialogTitle><DialogDescription>One shared Excel tracker that every computer opens and saves automatically.</DialogDescription></DialogHeader><OneDrivePanel drive={drive} busy={busy} hasWorkbook={false} onUse={useDriveFile} onCreate={putIntoDrive} onDisconnect={disconnect}/></DialogContent></Dialog></div>

  const context={state,commit,setModal,date}
  const queueReady=queue.filter((p)=>p.email&&!p.issues.includes("Duplicate party + invoice"))
  const confirmQueue=balanceParties.filter((p)=>!state.reminders.some((r)=>r.type==="Balance confirmation"&&r.date===date&&r.status!=="Cancelled"&&r.status!=="Failed"&&r.party===p.name))
  return <div className="debtors desk-shell"><Toaster position="top-right"/>{input}
    <aside className="desk-sidebar"><div className="desk-brand"><JalLogo/><span>Receivables desk</span></div><p className="nav-label">WORKSPACE</p><nav>{navigation.map(([key,label,Icon])=><button key={key} aria-label={label} title={label} className={page===key?"active":""} onClick={()=>go(key)}><Icon size={17}/><span>{label}</span>{key==="queue"&&<small>{queue.length}</small>}{key==="quality"&&issues.length+unmatched.length>0&&<i/>}</button>)}</nav><div className="sidebar-bottom"><div className="sidebar-file"><FileSpreadsheet size={22}/><div><b title={state.fileName}>{state.fileName}</b><span>{all.length.toLocaleString("en-IN")} invoices · {state.receipts.length} receipts</span></div></div><button className="sidebar-upload" disabled={busy} onClick={()=>fileInput.current?.click()}><Upload size={15}/> Open another workbook</button><div className="local-mode"><span/>Local workbook workspace</div></div></aside>
    <main className="desk-main"><header className="desk-header"><div><p className="breadcrumb">JAL Finance <span>/</span> {titles[page]}</p><h1>{titles[page]}</h1></div><div className="header-actions"><EmailBadge status={mail} senderName={senderName} onChangeName={askSenderName}/><span className="signed-in" title={`Signed in as ${account}`}><span>{account}</span><button onClick={()=>void signOut()}>Sign out</button></span><button className="desk-btn" aria-label="Open another workbook" title="Open another workbook" disabled={busy} onClick={()=>void openTrackerFile()}><Upload size={16}/></button><button className="desk-btn" aria-label="Email templates" title="Email templates" onClick={()=>setModal({kind:"settings"})}><Settings2 size={16}/><span>Email templates</span></button><button className="desk-btn primary" disabled={busy} onClick={()=>void save(false)}><Download size={16}/>{busy?"Working…":"Download Excel + logs"}</button></div></header>
    <div className="desk-content">{driveBase?<section className={`excel-save-panel onedrive ${driveSync.state==="conflict"||driveSync.state==="error"||dirty?"pending":"saved"}`} aria-live="polite"><div className="excel-file"><Cloud size={22}/><div><b title={drive.file?.name??state.fileName}>{drive.file?.name??state.fileName}</b><span>{driveLine(driveSync,dirty)}</span></div></div><div className="excel-save-actions">{driveSync.state==="conflict"?<><button className="desk-btn" onClick={()=>void resolveConflict("theirs")}>Open OneDrive version</button><button className="desk-btn" onClick={()=>void resolveConflict("mine")}>Keep mine and replace</button><button className="desk-btn" onClick={()=>void save(false)}><Download size={16}/>Download my copy</button></>:<><button className="desk-btn" onClick={()=>{go("logs");setLogTab("Activity");setSearch("")}}><History size={16}/>View log ({state.activity.length})</button>{driveSync.state==="error"&&<button className="desk-btn" onClick={()=>void pushToDrive(false)}>Retry save</button>}<button className="desk-btn" onClick={()=>setModal({kind:"onedrive"})}><Cloud size={16}/>OneDrive</button></>}</div><span className="recovery-note">Shared tracker in OneDrive{drive.account?` (${drive.account})`:""}. Changes save automatically, and other computers pick them up within about 30 seconds.</span></section>:<section className={`excel-save-panel ${fileConflict||pendingFile||dirty?"pending":"saved"}`} aria-live="polite"><div className="excel-file"><FileSpreadsheet size={22}/><div><b title={state.fileName}>{state.fileName}</b><span>{pendingFile?`Last time you worked in ${pendingFile.name}. Click Continue to keep saving into it.`:fileConflict?`${linkedName||"The Excel file"} was changed outside the app while the app also had unsaved changes. Choose which version to keep.`:autoSave?(dirty?`Saving into ${linkedName||"the Excel file"}…`:`All changes are saved into ${linkedName||"the Excel file"}${savedAt?` · ${new Date(savedAt).toLocaleString("en-IN",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"})}`:""}`):(dirty?"Latest changes are in this browser. Download Excel to keep them in your file.":"No pending changes. Logs are in your last saved or downloaded Excel file.")}</span></div></div><div className="excel-save-actions">{fileConflict?<><button className="desk-btn" onClick={()=>void resolveFileConflict("file")}>Open the file&apos;s version</button><button className="desk-btn" onClick={()=>void resolveFileConflict("mine")}>Keep the app&apos;s version</button><button className="desk-btn" onClick={()=>void save(false)}><Download size={16}/>Download a copy</button></>:<><button className="desk-btn" onClick={()=>{go("logs");setLogTab("Activity");setSearch("")}}><History size={16}/>View log ({state.activity.length})</button>{pendingFile?<button className="desk-btn primary" disabled={busy} onClick={()=>void continueWithFile()}><FileCheck2 size={16}/>Continue with {pendingFile.name}</button>:<button className="desk-btn" disabled={busy||autoSave} onClick={()=>{if(autoSaveAvailable)void save(true);else toast.info("For automatic Excel updates, open this app in Chrome or Edge and enable Excel auto-save. In this browser, use Download Excel + logs.")}}>{autoSave?<><FileCheck2 size={16}/>Excel auto-save on</>:<><FileSpreadsheet size={16}/>{autoSaveAvailable?"Enable Excel auto-save":"How to auto-save"}</>}</button>}</>}{drive.configured&&<button className="desk-btn" onClick={()=>setModal({kind:"onedrive"})}><Cloud size={16}/>{drive.connected?"Use OneDrive":"Connect OneDrive"}</button>}</div>{excelStatus&&<p className="excel-status-message">{excelStatus}</p>}<span className="recovery-note" title={draftStatus}>{draftStatus} · Activity Log and Reminder Log are included in every export.</span></section>}
      <div className="desk-filters"><label className="desk-search"><Search size={16}/><input aria-label="Search customer, invoice, contact" placeholder="Search customer, invoice, contact…" value={search} onChange={(e)=>setSearch(e.target.value)}/>{search&&<button aria-label="Clear search" onClick={()=>setSearch("")}><X size={14}/></button>}</label><select aria-label="Filter by area" value={area} onChange={(e)=>{setArea(e.target.value);setSelected(new Set())}}><option>All areas</option>{areas.map((a)=><option key={a}>{a}</option>)}</select><select aria-label="Filter by RSE" value={rse} onChange={(e)=>{setRse(e.target.value);setSelected(new Set())}}><option>All RSEs</option>{rses.map((a)=><option key={a}>{a}</option>)}</select><label className="review-date"><CalendarDays size={15}/><span>Review date</span><input aria-label="Review date" type="date" value={date} onChange={(e)=>{if(validDate(e.target.value)){setDate(e.target.value);setSelected(new Set())}}}/></label></div>
      {date!==today()&&<div className="desk-notice">Ageing and the queue use {displayDate(date)}. Balances include all receipts in the workbook; this is not a historical balance report.</div>}
      {page==="overview"&&<>
        <div className="overview-intro"><div><span className="eyebrow">COLLECTIONS AT A GLANCE</span><h2>Receivables at a glance</h2><p>{parties.length} customers with an open balance · {displayDate(date)}</p></div><button className="desk-btn primary" onClick={()=>{setDate(today());go("queue")}}><Bell size={17}/>Run today’s reminders <ArrowRight size={16}/></button></div>
        <div className="metric-grid"><Metric title="Total outstanding" value={shortMoney(total)} exact={money(total)} detail={`${open.length} open invoices`} icon={Wallet} onClick={()=>go("invoices")}/><Metric title="Overdue outstanding" value={shortMoney(overdue.reduce((s,i)=>s+i.outstanding,0))} exact={money(overdue.reduce((s,i)=>s+i.outstanding,0))} detail={`${overdue.length} invoices past due`} icon={Clock3} tone="amber" onClick={()=>{go("invoices");setInvoiceFilter("Overdue")}}/><Metric title="Critical exposure" value={shortMoney(critical.reduce((s,p)=>s+p.total,0))} detail={`${critical.length} critical customers`} icon={ShieldAlert} tone="red" onClick={()=>go("risk")}/><Metric title="Today’s queue" value={String(queue.length)} detail={`${queueReady.length} contacts ready to review`} icon={Bell} tone="teal" onClick={()=>go("queue")}/></div>
        <div className="overview-grid"><section className="desk-card"><CardHeader title="Overdue ageing" subtitle="Open balance by days past the due date" action="Explore ageing" onAction={()=>go("ageing")}/><AgeingBars invoices={open}/></section><CollectionsChart receipts={state.receipts.filter((r)=>filtered.some((i)=>allocationKey(i.party,i.number)===allocationKey(r.party,r.invoice)))} date={date}/></div>
        <div className="overview-grid wide-left"><section className="desk-card"><CardHeader title="Customers needing attention" subtitle="Prioritised by escalation level, then outstanding balance" action="View all" onAction={()=>go("risk")}/><PartyTable parties={parties.slice(0,5)} setModal={setModal} compact/></section><section className="desk-card"><CardHeader title="Your next actions" subtitle="Keep the follow-up cycle moving"/><div className="next-actions"><ActionRow icon={Bell} title="Review reminder queue" detail={`${queue.length} customers eligible for follow-up`} onClick={()=>go("queue")}/><ActionRow icon={Users} title="Confirm customer balances" detail={`${parties.length} open customer accounts`} onClick={()=>go("balances")}/><ActionRow icon={CircleAlert} title="Resolve data exceptions" detail={`${issues.length} invoice rows · ${unmatched.length} unmatched receipts`} onClick={()=>go("quality")}/><ActionRow icon={Download} title="Save your updated workbook" detail={dirty?"Export your latest changes and history":"Keep one current file for the next session"} onClick={()=>void save(false)}/></div></section></div>
        <section className="desk-card"><CardHeader title="Outstanding by sales owner" subtitle="Top RSEs by open balance" action="Invoice register" onAction={()=>go("invoices")}/><OwnerBars invoices={open} field="rse"/></section>
      </>}
      {page==="queue"&&<><div className="page-intro"><div><h2>Payment reminders</h2><p>Due in the next {state.settings.preDueDays} days or overdue · repeat follow-up after {state.settings.repeatDays} days. Held invoices and drafts already prepared today are excluded.</p></div><button className="desk-btn" onClick={()=>{setDate(today());setSelected(new Set());toast.info("Queue refreshed for today")}}><CalendarDays size={16}/>Run today’s reminders</button></div><div className={`desk-notice${mail.configured&&mail.mode==="live"?" live":""}`}><Mail size={17}/><div>{sendingNote(mail)} Preparing a draft does not change “Last Reminder Sent”.</div></div><QueueTable parties={queue} selected={selected} toggle={toggle} setSelected={setSelected} setModal={setModal} type="Reminder"/><div className="selection-bar"><div><b>{queue.filter((p)=>selected.has(p.key)).length} customers selected</b><span>{money(queue.filter((p)=>selected.has(p.key)).reduce((s,p)=>s+p.total,0))} outstanding</span></div><button className="desk-btn primary" disabled={!queue.some((p)=>selected.has(p.key))} onClick={()=>prepare(queue,"Reminder")}><FileCheck2 size={17}/>Prepare selected drafts</button><SendButton drafts={state.reminders.filter((r)=>r.type==="Reminder"&&r.status==="Prepared"&&r.date===date)} mail={mail} sending={sending} onSend={sendDrafts}/></div></>}
      {page==="balances"&&<><div className="page-intro"><div><h2>Reconcile the balance with your customers.</h2><p>Confirm the whole open balance per customer. Record a confirmation or dispute without changing invoice amounts.</p></div><button className="desk-btn primary" disabled={!confirmQueue.some((p)=>selected.has(p.key))} onClick={()=>prepare(confirmQueue,"Balance confirmation")}><Mail size={16}/>Prepare selected confirmations</button><SendButton drafts={state.reminders.filter((r)=>r.type==="Balance confirmation"&&r.status==="Prepared"&&r.date===date)} mail={mail} sending={sending} onSend={sendDrafts}/></div><div className="desk-card"><CardHeader title="Open customer balances" subtitle="Selection prepares confirmation requests; a recorded response is kept separately."/><PagedRows rows={balanceParties} render={(rows)=><div className="desk-table-wrap"><table className="desk-table"><thead><tr><th><input type="checkbox" aria-label="Select all ready balance confirmations" checked={confirmQueue.some((p)=>Boolean(p.email))&&confirmQueue.filter((p)=>p.email&&!p.issues.includes("Duplicate party + invoice")).every((p)=>selected.has(p.key))} onChange={(e)=>setSelected(e.target.checked?new Set(confirmQueue.filter((p)=>p.email&&!p.issues.includes("Duplicate party + invoice")).map((p)=>p.key)):new Set())}/></th><th>Customer</th><th className="numeric">Open balance</th><th>Last response</th><th>Request</th><th>Actions</th></tr></thead><tbody>{rows.map((p)=>{const last=state.confirmations.filter((c)=>partyKey(c.party)===p.key).at(-1);const prepared=!confirmQueue.some((c)=>c.key===p.key);return <tr key={p.key}><td><input aria-label={`Select ${p.name}`} type="checkbox" checked={selected.has(p.key)} disabled={!p.email||prepared||p.issues.includes("Duplicate party + invoice")} onChange={()=>toggle(p.key)}/></td><td><b>{p.name}</b><small>{p.invoices.length} open invoices · {p.email||"Contact needs review"}</small></td><td className="numeric">{money(p.total)}</td><td>{last?<><Status value={last.status}/><small>{displayDate(last.date)} · {money(last.amount)}</small></>:<span className="muted">No response recorded</span>}</td><td>{prepared?<Status value="Prepared today"/>:<button className="text-action" onClick={()=>setModal({kind:"message",party:p,type:"Balance confirmation"})}>Preview request</button>}</td><td><div className="row-actions"><button onClick={()=>setModal({kind:"confirmation",party:p.key})}>Record response</button><button onClick={()=>setModal({kind:"contact",party:p.name})}>Contact</button></div></td></tr>})}</tbody></table></div>}/></div></>}
      {page==="invoices"&&<><div className="page-intro"><div><h2>Invoices and payments</h2><p>Open an invoice for its timeline, payments, contact and follow-up actions.</p></div><select aria-label="Invoice status" value={invoiceFilter} onChange={(e)=>setInvoiceFilter(e.target.value)}>{["All invoices","Open","Overdue","Part paid","Paid","Flagged","On hold"].map((f)=><option key={f}>{f}</option>)}</select></div><InvoiceTable invoices={filtered.filter((i)=>invoiceFilter==="All invoices"||invoiceFilter==="Open"&&i.outstanding>0||invoiceFilter==="Overdue"&&i.overdue>0||invoiceFilter===i.status||invoiceFilter==="Flagged"&&i.followup.flag||invoiceFilter==="On hold"&&i.followup.hold)} setModal={setModal}/></>}
      {page==="ageing"&&<><div className="page-intro"><div><h2>Know how long your balance has been open.</h2><p>Overdue ageing uses the due date. Invoice ageing uses the invoice date.</p></div></div><div className="overview-grid"><section className="desk-card"><CardHeader title="Days overdue" subtitle="Only open balances; not due includes invoices due today"/><AgeingBars invoices={open}/></section><section className="desk-card"><CardHeader title="Age from invoice date" subtitle="Open balances grouped by time since invoicing"/><AgeingBars invoices={open} invoiceAge/></section></div><div className="overview-grid"><section className="desk-card"><CardHeader title="Area exposure" subtitle="Top areas by outstanding balance"/><OwnerBars invoices={open} field="area"/></section><section className="desk-card"><CardHeader title="Upcoming due dates" subtitle="Next 7 days, based on the review date"/><div className="upcoming-list">{Array.from({length:7},(_,index)=>{const d=new Date((dayNumber(date)+index)*86400000).toISOString().slice(0,10),rows=open.filter((i)=>i.due===d);return <div key={d}><span>{displayDate(d)}</span><small>{rows.length} invoices</small><b>{money(rows.reduce((s,i)=>s+i.outstanding,0))}</b></div>})}</div></section></div><InvoiceTable invoices={overdue.sort((a,b)=>b.overdue-a.overdue)} setModal={setModal}/></>}
      {page==="risk"&&<><div className="page-intro"><div><h2>Customers to follow up</h2><p>Critical: 45+ days overdue or Management / Legal escalation. High: 30–44 days, CFO escalation or a manual flag. Watch: 15–29 days.</p></div></div><div className="risk-grid">{["Critical","High","Watch","Standard"].map((risk)=>{const rows=parties.filter((p)=>p.risk===risk);return <div className="risk-card" key={risk}><Status value={risk}/><strong>{shortMoney(rows.reduce((s,p)=>s+p.total,0))}</strong><p>{rows.length} customers</p></div>})}</div><section className="desk-card"><CardHeader title="Debtor exposure" subtitle="Open the oldest invoice to record a flag, hold or escalation. These controls do not overwrite formula columns."/><PartyTable parties={parties} setModal={setModal}/></section></>}
      {page==="logs"&&<><div className="page-intro"><div><h2>Excel logs and history</h2><p>Every edit is recorded here. Download Excel or enable Excel auto-save to write these entries into the file.</p></div><div className="segmented">{["Activity","Reminder drafts","Confirmations"].map((tab)=><button key={tab} className={logTab===tab?"active":""} onClick={()=>setLogTab(tab)}>{tab}</button>)}</div></div><div className="log-summary"><div><b>{state.activity.length}</b><span>Activity entries</span></div><div><b>{state.reminders.length}</b><span>Message records</span></div><div><b>{state.confirmations.length}</b><span>Balance responses</span></div><p>{autoSave?"Excel auto-save is connected.":dirty?"These entries have not yet been written to your Excel file.":"Log sheets are in your last exported Excel copy."} Sheets: Activity Log, Reminder Log and Balance Confirmations.</p></div>{logTab==="Activity"?<ActivityTable activity={[...state.activity].reverse().filter((a)=>!search||[a.party,a.invoice,a.action,a.detail].some((v)=>v.toLowerCase().includes(search.toLowerCase())))}/>:logTab==="Reminder drafts"?<section className="desk-card"><PagedRows rows={[...state.reminders].reverse().filter((r)=>!search||[r.party,r.email,r.subject,...r.invoiceNumbers].some((v)=>v.toLowerCase().includes(search.toLowerCase())))} render={(rows)=><div className="desk-table-wrap"><table className="desk-table"><thead><tr><th>Created</th><th>Customer / type</th><th>Recipient</th><th className="numeric">Amount</th><th>Status</th><th>Message</th></tr></thead><tbody>{rows.map((r)=><tr key={r.id}><td>{new Date(r.at).toLocaleString("en-IN")}</td><td><b>{r.party}</b><small>{r.type} · {r.invoiceNumbers.length} invoices</small></td><td>{r.email}</td><td className="numeric">{money(r.amount)}</td><td><Status value={r.status}/>{r.result&&<small title={r.result}>{r.result}</small>}</td><td><button className="text-action" onClick={()=>setModal({kind:"draft",id:r.id})}>View message</button></td></tr>)}</tbody></table></div>}/></section>:<ConfirmationTable confirmations={[...state.confirmations].reverse()}/>}<section className="desk-card legacy-history"><CardHeader title="Last Reminder Sent (Excel column Q)" subtitle="The tracker’s own record. Earlier entries are kept as they were; each live email sent from the app adds an entry such as “24.10.26 MAIL”."/><PagedRows rows={filtered.filter((i)=>i.legacyReminder)} render={(rows)=><div className="desk-table-wrap"><table className="desk-table"><thead><tr><th>Customer</th><th>Invoice</th><th>Original reminder entry</th></tr></thead><tbody>{rows.map((i)=><tr key={i.id}><td>{i.party}</td><td>{i.number}</td><td>{i.legacyReminder}</td></tr>)}</tbody></table></div>}/></section></>}
      {page==="quality"&&<><div className="page-intro"><div><h2>Resolve exceptions before following up.</h2><p>Unmatched receipts are not silently allocated. Credit entries and overpayments stay visible and never reduce another invoice’s balance.</p></div></div>{state.importNotes.map((note)=><div className="desk-notice amber" key={note}><CircleAlert size={18}/>{note}</div>)}<div className="quality-metrics"><span><b>{all.filter((i)=>i.outstanding>0&&!validEmail(i.email)).length}</b> open rows missing valid email</span><span><b>{unmatched.length}</b> unmatched receipts ({money(unmatched.reduce((s,r)=>s+r.amount,0))})</span><span><b>{all.filter((i)=>i.issues.includes("Receipt exceeds invoice")).length}</b> overpaid invoices</span><span><b>{all.filter((i)=>i.amount<0).length}</b> credit entries</span></div><section className="desk-card"><CardHeader title="Invoice exceptions" subtitle="Source row numbers refer to Debtors Data in your workbook"/><PagedRows rows={filtered.filter((i)=>i.issues.length)} render={(rows)=><div className="desk-table-wrap"><table className="desk-table"><thead><tr><th>Excel row</th><th>Customer / invoice</th><th>Exception</th><th className="numeric">Outstanding</th><th>Action</th></tr></thead><tbody>{rows.map((i)=><tr key={i.id}><td>{i.row}</td><td><b>{i.party}</b><small>{i.number}</small></td><td>{i.issues.join(" · ")}</td><td className="numeric">{money(i.outstanding)}</td><td><button className="text-action" onClick={()=>setModal({kind:"invoice",id:i.id})}>Review</button></td></tr>)}</tbody></table></div>}/></section><section className="desk-card"><CardHeader title="Unmatched receipts" subtitle="These party + invoice pairs do not match the invoice register. Correct the source workbook after checking the accounting record."/><PagedRows rows={unmatched} render={(rows)=><div className="desk-table-wrap"><table className="desk-table"><thead><tr><th>Receipts row</th><th>Party</th><th>Invoice</th><th>Date</th><th className="numeric">Amount</th></tr></thead><tbody>{rows.map((r)=><tr key={r.id}><td>{r.row}</td><td>{r.party}</td><td>{r.invoice}</td><td>{displayDate(r.date)}</td><td className="numeric">{money(r.amount)}</td></tr>)}</tbody></table></div>}/></section></>}
      <footer className="desk-footer"><span>JAL Receivables Desk</span><span>Balances from all receipts · {filtered.length.toLocaleString("en-IN")} of {all.length.toLocaleString("en-IN")} invoices in view</span><span>Save the updated workbook before sharing or switching devices.</span></footer>
    </div></main>
    <Dialog open={modal!==null} onOpenChange={(open)=>{if(!open)setModal(null)}}><DialogContent className="debtors desk-dialog max-h-[92vh] overflow-y-auto sm:max-w-3xl" showCloseButton={false}><button className="dialog-close-action" aria-label="Close" onClick={()=>setModal(null)}><X size={20}/></button><DialogHeader><DialogTitle>{modal?.kind==="invoice"?"Invoice detail":modal?.kind==="payment"?modal.receipt?"Update payment record":"Record payment":modal?.kind==="contact"?"Update customer contact":modal?.kind==="followup"?"Flag or escalate invoice":modal?.kind==="confirmation"?"Record balance response":modal?.kind==="settings"?"Email templates & reminder rules":modal?.kind==="onedrive"?"OneDrive tracker":"Review message"}</DialogTitle><DialogDescription>{modal?.kind==="settings"?"Customise reminder and balance confirmation messages before preparing drafts.":modal?.kind==="message"||modal?.kind==="draft"?"Review the recipient, subject and complete message before sending.":modal?.kind==="onedrive"?"One shared Excel tracker that every computer opens and saves automatically.":"Changes are kept in your browser draft and included when you save the workbook."}</DialogDescription></DialogHeader>
      {modal?.kind==="invoice"&&all.find((i)=>i.id===modal.id)&&<InvoiceDetail invoice={all.find((i)=>i.id===modal.id)!} {...context}/>}
      {modal?.kind==="contact"&&<ContactForm key={modal.party} party={modal.party} {...context}/>}
      {modal?.kind==="payment"&&<PaymentForm key={`${modal.invoice}-${modal.receipt}`} invoice={all.find((i)=>i.id===modal.invoice)!} receipt={state.receipts.find((r)=>r.id===modal.receipt)??(all.find((i)=>i.id===modal.invoice)!.outstanding<=0?[...state.receipts].reverse().find((r)=>allocationKey(r.party,r.invoice)===allocationKey(all.find((i)=>i.id===modal.invoice)!.party,all.find((i)=>i.id===modal.invoice)!.number)):undefined)} {...context}/>}
      {modal?.kind==="followup"&&<FollowupForm key={modal.id} invoice={all.find((i)=>i.id===modal.id)!} {...context}/>}
      {modal?.kind==="confirmation"&&<ConfirmationForm key={modal.party} party={groupParties(all.filter((i)=>i.outstanding>0)).find((p)=>p.key===modal.party)!} {...context}/>}
      {modal?.kind==="settings"&&<SettingsForm key="settings" {...context}/>}
      {modal?.kind==="onedrive"&&<OneDrivePanel drive={drive} busy={busy} hasWorkbook onUse={useDriveFile} onCreate={putIntoDrive} onDisconnect={disconnect}/>}
      {modal?.kind==="message"&&<><MessagePreview note={sendingNote(mail)} message={{...renderMessage(state.settings,modal.party,date,modal.type),email:modal.party.email}} onContact={()=>setModal({kind:"contact",party:modal.party.name})}/><button className="desk-btn primary" disabled={!modal.party.email||modal.party.issues.includes("Duplicate party + invoice")} onClick={()=>prepare([modal.party],modal.type,true)}><FileCheck2 size={17}/>Prepare this draft</button></>}
      {modal?.kind==="draft"&&(()=>{const draft=state.reminders.find((r)=>r.id===modal.id)!;return <><MessagePreview note={draft.result??sendingNote(mail)} message={draft}/><div className="draft-controls"><Status value={draft.status}/>{draft.sentAt&&<span className="muted">Sent {new Date(draft.sentAt).toLocaleString("en-IN")} to {draft.sentTo}{draft.sentBy&&` by ${draft.sentBy}`}</span>}{draft.status==="Prepared"&&<><button className="desk-btn" onClick={()=>cancelDraft(draft)}>Cancel draft</button><SendButton single drafts={[draft]} mail={mail} sending={sending} onSend={sendDrafts}/></>}</div></>})()}
    </DialogContent></Dialog>
  </div>
}

function Metric({title,value,detail,exact,icon:Icon,tone="navy",onClick}:{title:string;value:string;detail:string;exact?:string;icon:typeof Wallet;tone?:string;onClick:()=>void}) {
  return <button className={`metric-card ${tone}`} onClick={onClick}><div><span>{title}</span><Icon size={17}/></div><strong title={exact}>{value}</strong><p>{detail}<ArrowUpRight size={15}/></p></button>
}
function CardHeader({title,subtitle,action,onAction}:{title:string;subtitle?:string;action?:string;onAction?:()=>void}) {return <div className="card-heading"><div><h3>{title}</h3>{subtitle&&<p>{subtitle}</p>}</div>{action&&<button className="text-action" onClick={onAction}>{action}<ArrowRight size={14}/></button>}</div>}
function Status({value}:{value:string}) {const tone=/critical|failed|disputed|legal|management/i.test(value)?"red":/high|overdue|cfo|watch|part paid|hold/i.test(value)?"amber":/paid|confirmed|accepted|closed/i.test(value)?"teal":"slate";return <span className={`desk-status ${tone}`}>{value}</span>}
function ActionRow({icon:Icon,title,detail,onClick}:{icon:typeof Bell;title:string;detail:string;onClick:()=>void}) {return <button onClick={onClick}><span><Icon size={17}/></span><div><b>{title}</b><p>{detail}</p></div><ArrowRight size={15}/></button>}
function Empty({text="No records in this view"}:{text?:string}) {return <div className="desk-empty"><FileCheck2 size={28}/><b>{text}</b><p>Try another filter or review the workbook data.</p></div>}
function PagedRows<T>({rows,render}:{rows:T[];render:(rows:T[])=>React.ReactNode}) {
  const [page,setPage]=useState(0),size=25,last=Math.max(0,Math.ceil(rows.length/size)-1),current=Math.min(page,last)
  if(!rows.length)return <Empty/>
  return <>{render(rows.slice(current*size,(current+1)*size))}<div className="table-pagination"><span>{current*size+1}–{Math.min((current+1)*size,rows.length)} of {rows.length.toLocaleString("en-IN")} records</span><div><button disabled={current===0} onClick={()=>setPage(current-1)}>Previous</button><span>{current+1} / {last+1}</span><button disabled={current===last} onClick={()=>setPage(current+1)}>Next</button></div></div></>
}
function AgeingBars({invoices,invoiceAge=false}:{invoices:Invoice[];invoiceAge?:boolean}) {
  const labels=invoiceAge?["0–30 days","31–45 days","46–60 days","61–90 days","90+ days"]:ageingLabels
  const rows=labels.map((label,index)=>{const records=invoices.filter((i)=>invoiceAge?(i.age<=30?0:i.age<=45?1:i.age<=60?2:i.age<=90?3:4)===index:i.ageing===label);return {label,amount:records.reduce((s,i)=>s+i.outstanding,0),count:records.length}}),total=rows.reduce((s,r)=>s+r.amount,0)
  return <div className="ageing-bars">{rows.map((r,index)=><div className="ageing-bar-row" key={r.label}><span>{r.label}<small>{r.count} invoices</small></span><div><i style={{width:`${total?r.amount/Math.max(...rows.map((v)=>v.amount))*100:0}%`,background:["#92a8bf","#5d7e9e","#d3aa59","#d28a55","#c36955","#a7474c"][index]}}/></div><b title={money(r.amount)}>{shortMoney(r.amount)}</b></div>)}<div className="chart-total"><span>Total open balance</span><b>{money(total)}</b></div></div>
}
function OwnerBars({invoices,field}:{invoices:Invoice[];field:"rse"|"area"}) {
  const groups=new Map<string,number>();for(const i of invoices)groups.set(i[field]||"Unassigned",(groups.get(i[field]||"Unassigned")??0)+i.outstanding)
  const rows=[...groups].sort((a,b)=>b[1]-a[1]).slice(0,6),max=rows[0]?.[1]??1
  return <div className="owner-bars">{rows.map(([name,amount])=><div key={name}><div><span title={name}>{name}</span><b>{shortMoney(amount)}</b></div><div className="owner-track"><i style={{width:`${amount/max*100}%`}}/></div></div>)}{!rows.length&&<Empty/>}</div>
}
function CollectionsChart({receipts,date}:{receipts:Receipt[];date:string}) {
  const days=Array.from({length:14},(_,index)=>{const d=new Date((dayNumber(date)-13+index)*86400000).toISOString().slice(0,10);return {date:d,total:receipts.filter((r)=>r.date===d).reduce((s,r)=>s+r.amount,0)}}),total=days.reduce((s,d)=>s+d.total,0),max=Math.max(1,...days.map((d)=>d.total))
  return <section className="desk-card"><CardHeader title="Recent collections" subtitle="Matched receipts over the last 14 days"/><div className="collection-head"><strong>{shortMoney(total)}</strong><span><ArrowDownLeft size={14}/> allocated collections</span></div><div className="collection-chart">{days.map((d)=><div key={d.date} title={`${displayDate(d.date)}: ${money(d.total)}`}><i style={{height:`${Math.max(2,d.total/max*125)}px`}}/><small>{d.date.slice(-2)}</small></div>)}</div><div className="collection-range"><span>{displayDate(days[0].date)}</span><span>{displayDate(date)}</span></div></section>
}
function InvoiceTable({invoices,setModal}:{invoices:Invoice[];setModal:(modal:Modal)=>void}) {
  return <section className="desk-card"><PagedRows rows={invoices} render={(rows)=><div className="desk-table-wrap"><table className="desk-table"><thead><tr><th>Invoice / customer</th><th>Invoice date</th><th>Due date</th><th className="numeric">Invoice total</th><th className="numeric">Outstanding</th><th>Status</th><th>Follow-up</th><th>Actions</th></tr></thead><tbody>{rows.map((i)=><tr key={i.id}><td><button className="invoice-link" onClick={()=>setModal({kind:"invoice",id:i.id})}>{i.number}</button><small title={i.party}>{i.party}</small><button className="contact-shortcut" onClick={()=>setModal({kind:"contact",party:i.party})}>Edit contact</button></td><td>{displayDate(i.issued)}</td><td>{displayDate(i.due)}<small>{i.overdue>0?`${i.overdue} days overdue`:i.outstanding<=0?"Closed":i.daysToDue===0?"Due today":`${i.daysToDue} days to due`}</small></td><td className="numeric">{money(i.amount)}</td><td className="numeric"><b>{money(i.outstanding)}</b></td><td><Status value={i.status}/></td><td><div className="followup-tags">{i.followup.flag&&<Flag size={13}/>}<span>{i.followup.hold?"On hold":i.followup.escalation||i.stage}</span></div></td><td><div className="row-actions primary-row-actions"><button className="desk-btn" disabled={i.amount<=0||i.issues.includes("Duplicate party + invoice")} onClick={()=>setModal({kind:"payment",invoice:i.id})}>{i.outstanding>0?"Record payment":"Update payment"}</button><button className="desk-btn" onClick={()=>setModal({kind:"followup",id:i.id})}>Follow-up</button></div></td></tr>)}</tbody></table></div>}/></section>
}
function PartyTable({parties,setModal,compact=false}:{parties:Party[];setModal:(modal:Modal)=>void;compact?:boolean}) {
  const render=(rows:Party[]) => <div className="desk-table-wrap"><table className="desk-table"><thead><tr><th>Customer</th><th className="numeric">Outstanding</th><th>Oldest overdue</th><th>Risk</th><th>Action</th></tr></thead><tbody>{rows.map((p)=><tr key={p.key}><td><b>{p.name}</b><small>{p.invoices.length} invoices{!compact&&` · ${p.email||"Contact needs review"}`}</small></td><td className="numeric">{money(p.total)}</td><td>{p.maxOverdue} days</td><td><Status value={p.risk}/></td><td><button className="text-action" onClick={()=>setModal({kind:"invoice",id:[...p.invoices].sort((a,b)=>b.overdue-a.overdue)[0].id})}>Details <ArrowUpRight size={13}/></button><button className="desk-btn" onClick={()=>setModal({kind:"followup",id:[...p.invoices].sort((a,b)=>b.overdue-a.overdue)[0].id})}>Follow-up</button></td></tr>)}</tbody></table></div>
  return compact?parties.length?render(parties):<Empty/>:<PagedRows rows={parties} render={render}/>
}
function QueueTable({parties,selected,toggle,setSelected,setModal,type}:{parties:Party[];selected:Set<string>;toggle:(key:string)=>void;setSelected:(value:Set<string>)=>void;setModal:(modal:Modal)=>void;type:Reminder["type"]}) {
  const ready=parties.filter((p)=>p.email&&!p.issues.includes("Duplicate party + invoice"))
  return <section className="desk-card"><div className="queue-heading"><p><b>{parties.length}</b> customers · <b>{ready.length}</b> ready for review</p><div><button className="text-action" onClick={()=>setSelected(new Set())}>Clear</button><button className="desk-btn" onClick={()=>setSelected(new Set(ready.map((p)=>p.key)))}>Select all ready</button></div></div><PagedRows rows={parties} render={(rows)=><div className="desk-table-wrap"><table className="desk-table"><thead><tr><th>Select</th><th>Customer / contact</th><th>Invoices</th><th className="numeric">Outstanding</th><th>Priority / stage</th><th>Last reminder</th><th>Review</th></tr></thead><tbody>{rows.map((p)=><tr key={p.key}><td><input type="checkbox" aria-label={`Select ${p.name}`} checked={selected.has(p.key)} disabled={!p.email||p.issues.includes("Duplicate party + invoice")} onChange={()=>toggle(p.key)}/></td><td><b>{p.name}</b><button className={`contact-link ${p.email?"":"needs-contact"}`} onClick={()=>setModal({kind:"contact",party:p.name})}>{p.email||"Update contact"}</button></td><td>{p.invoices.length}</td><td className="numeric">{money(p.total)}</td><td><Status value={p.risk}/><small>{[...p.invoices].sort((a,b)=>b.overdue-a.overdue)[0].stage}</small></td><td>{[...p.invoices].map((i)=>i.lastReminder).sort().at(-1)?displayDate([...p.invoices].map((i)=>i.lastReminder).sort().at(-1)!):"No date recorded"}</td><td><button className="text-action" onClick={()=>setModal({kind:"message",party:p,type})}>Preview <ArrowUpRight size={13}/></button></td></tr>)}</tbody></table></div>}/></section>
}
function ActivityTable({activity}:{activity:WorkbookState["activity"]}) {return <section className="desk-card"><PagedRows rows={activity} render={(rows)=><div className="desk-table-wrap"><table className="desk-table"><thead><tr><th>Time</th><th>Action</th><th>Customer / invoice</th><th>Details</th></tr></thead><tbody>{rows.map((a)=><tr key={a.id}><td>{new Date(a.at).toLocaleString("en-IN")}</td><td><b>{a.action}</b></td><td>{a.party||"Workbook"}<small>{a.invoice}</small></td><td>{a.detail}</td></tr>)}</tbody></table></div>}/></section>}
function ConfirmationTable({confirmations}:{confirmations:Confirmation[]}) {return <section className="desk-card"><PagedRows rows={confirmations} render={(rows)=><div className="desk-table-wrap"><table className="desk-table"><thead><tr><th>Date</th><th>Customer</th><th className="numeric">Balance snapshot</th><th>Response</th><th>Reference / note</th></tr></thead><tbody>{rows.map((c)=><tr key={c.id}><td>{displayDate(c.date)}</td><td>{c.party}</td><td className="numeric">{money(c.amount)}</td><td><Status value={c.status}/></td><td>{c.reference}<small>{c.note}</small></td></tr>)}</tbody></table></div>}/></section>}

type FormContext={state:WorkbookState;commit:(next:WorkbookState)=>void;setModal:(modal:Modal)=>void;date:string}
function InvoiceDetail({invoice:i,state,setModal}:FormContext&{invoice:Invoice}) {
  const receipts=state.receipts.filter((r)=>allocationKey(r.party,r.invoice)===allocationKey(i.party,i.number)),events=state.activity.filter((a)=>a.invoiceId===i.id||!a.invoiceId&&partyKey(a.party)===partyKey(i.party)),importedReceipts=receipts.filter((r)=>!events.some((a)=>a.action==="Payment recorded"&&a.after.includes(r.id)))
  return <div className="invoice-detail"><div className="detail-heading"><div><span className="eyebrow">{i.number}</span><h2>{i.party}</h2><p>{i.area} · {i.rse}</p></div><Status value={i.status}/></div><div className="detail-metrics"><div><span>Invoice total</span><b>{money(i.amount)}</b></div><div><span>Received</span><b>{money(i.received)}</b></div><div><span>Outstanding</span><b>{money(i.outstanding)}</b></div></div><div className="detail-info"><p><span>Invoice date</span><b>{displayDate(i.issued)}</b></p><p><span>Due date</span><b>{displayDate(i.due)}</b></p><p><span>Days overdue</span><b>{i.overdue}</b></p><p><span>Source row</span><b>Debtors Data · {i.row}</b></p><p><span>Contact</span><button className="text-action" onClick={()=>setModal({kind:"contact",party:i.party})}>{i.email||"Add email"}</button></p><p><span>Follow-up</span><b>{i.followup.hold?"On hold":i.followup.escalation||i.stage}</b></p></div>{i.issues.length>0&&<div className="desk-notice amber">{i.issues.join(" · ")}</div>}<div className="detail-actions"><button className="desk-btn primary" disabled={i.outstanding<=0||i.issues.includes("Duplicate party + invoice")} onClick={()=>setModal({kind:"payment",invoice:i.id})}><Wallet size={16}/>Record payment</button><button className="desk-btn" onClick={()=>setModal({kind:"followup",id:i.id})}><Flag size={16}/>Flag / escalate / notes</button></div>{i.followup.note&&<div className="detail-note"><b>Follow-up note</b><p>{i.followup.note}</p></div>}{i.remarks&&<div className="detail-note"><b>Sales / RSE remarks</b><p>{i.remarks}</p></div>}<h3>Payment records</h3>{receipts.length?<div className="payment-records">{receipts.map((r)=><div key={r.id}><span><b>{money(r.amount)}</b><small>{displayDate(r.date)} · {r.reference||"No reference"}</small>{r.remarks&&<small>{r.remarks}</small>}</span><button className="text-action" onClick={()=>setModal({kind:"payment",invoice:i.id,receipt:r.id})}>Update payment</button></div>)}</div>:<p className="muted">No receipts allocated to this invoice.</p>}<h3>Invoice timeline</h3><div className="invoice-timeline"><TimelineItem title="Invoice issued" date={displayDate(i.issued)} detail={money(i.amount)}/><TimelineItem title="Payment due" date={displayDate(i.due)} detail={i.stage}/>{i.legacyReminder&&<TimelineItem title="Imported reminder history" date="Original workbook entry" detail={i.legacyReminder}/>} {[...importedReceipts.map((r)=>({at:r.date,title:"Imported collection",detail:`${money(r.amount)} · ${r.reference||"No reference"}`})),...events.map((a)=>({at:a.at,title:a.action,detail:a.detail}))].sort((a,b)=>a.at.localeCompare(b.at)).map((e,index)=><TimelineItem key={`${e.at}-${index}`} title={e.title} date={displayDate(e.at.slice(0,10))} detail={e.detail}/>)}</div></div>
}
function TimelineItem({title,date,detail}:{title:string;date:string;detail:string}) {return <div><i/><span>{date}</span><b>{title}</b><p>{detail}</p></div>}
function ContactForm({party,state,commit,setModal}:FormContext&{party:string}) {
  const rows=state.invoices.filter((i)=>partyKey(i.party)===partyKey(party)),emails=[...new Set(rows.map((i)=>i.email).filter(Boolean))],[email,setEmail]=useState(emails.length===1?emails[0]:"")
  return <form className="desk-form" onSubmit={(e)=>{e.preventDefault();try{commit(updateContact(state,party,email));setModal(null);toast.success("Contact updated for all invoices belonging to this customer")}catch(error){toast.error((error as Error).message)}}}><h3>{party}</h3><p className="muted">This updates the Email ID on all {rows.length} invoice rows for this customer. Previously prepared drafts keep their original recipient snapshot.</p>{emails.length>1&&<div className="desk-notice amber">Conflicting contacts: {emails.join(" · ")}. Enter the intended recipient below.</div>}<Field label="Email address(es)"><input required value={email} onChange={(e)=>setEmail(e.target.value)} placeholder="accounts@example.com; owner@example.com"/></Field><button className="desk-btn primary" type="submit">Save contact</button></form>
}
function PaymentForm({invoice,receipt,state,commit,setModal}:FormContext&{invoice:Invoice;receipt?:Receipt}) {
  const [paymentDate,setPaymentDate]=useState(receipt?.date||today()),[amount,setAmount]=useState(String(receipt?.amount??"")),[reference,setReference]=useState(receipt?.reference??""),[note,setNote]=useState(receipt?.remarks??"")
  return <form className="desk-form" onSubmit={(e)=>{e.preventDefault();try{commit(recordReceipt(state,invoice,{date:paymentDate,amount:Number(amount),reference:reference.trim(),remarks:note},receipt?.id));setModal(null);toast.success(receipt?"Payment corrected and audit history kept":"Collection recorded; balance recalculated")}catch(error){toast.error((error as Error).message)}}}><h3>{invoice.number} · {invoice.party}</h3><Field label="Payment record"><select aria-label="Payment record" value={receipt?.id??""} onChange={(e)=>setModal({kind:"payment",invoice:invoice.id,receipt:e.target.value||undefined})}><option value="" disabled={invoice.outstanding<=0}>New payment</option>{state.receipts.filter((r)=>allocationKey(r.party,r.invoice)===allocationKey(invoice.party,invoice.number)).map((r)=><option value={r.id} key={r.id}>{displayDate(r.date)} · {money(r.amount)} · {r.reference||"Existing receipt"}</option>)}</select></Field><div className="desk-notice">Outstanding: {money(invoice.outstanding)}. {receipt?"The previous receipt amount is included when checking this correction.":"Part payments are supported. A receipt is allocated only to this invoice."}</div><div className="form-two"><Field label="Payment date"><input required type="date" value={paymentDate} onChange={(e)=>setPaymentDate(e.target.value)}/></Field><Field label="Amount received (INR)"><input required type="number" min="0.01" step="0.01" value={amount} onChange={(e)=>setAmount(e.target.value)}/></Field></div><Field label="Bank / UTR reference"><input value={reference} onChange={(e)=>setReference(e.target.value)} placeholder="Transaction or receipt reference"/></Field><Field label="Remarks"><textarea value={note} onChange={(e)=>setNote(e.target.value)} rows={3}/></Field><button className="desk-btn primary" type="submit">{receipt?"Update payment record":"Record payment"}</button></form>
}
function FollowupForm({invoice,state,commit,setModal}:FormContext&{invoice:Invoice}) {
  const [flag,setFlag]=useState(invoice.followup.flag),[hold,setHold]=useState(invoice.followup.hold),[escalation,setEscalation]=useState(invoice.followup.escalation),[note,setNote]=useState(invoice.followup.note),[remarks,setRemarks]=useState(invoice.remarks)
  return <form className="desk-form" onSubmit={(e)=>{e.preventDefault();const next={flag,hold,escalation,note};commit({...state,revision:state.revision+1,invoices:state.invoices.map((i)=>i.id===invoice.id?{...i,remarks}:i),followups:{...state.followups,[invoice.id]:next},activity:[...state.activity,addActivity(state,"Follow-up updated",invoice,`${flag?"Flagged · ":""}${hold?"On hold · ":""}${escalation||"Automatic escalation"}`,JSON.stringify({followup:invoice.followup,remarks:invoice.remarks}),JSON.stringify({followup:next,remarks}))]});setModal(null);toast.success("Follow-up saved")} }><h3>{invoice.number} · {invoice.party}</h3><div className="check-field"><input id="flag-invoice" type="checkbox" checked={flag} onChange={(e)=>setFlag(e.target.checked)}/><label htmlFor="flag-invoice">Flag for priority review</label></div><div className="check-field"><input id="hold-invoice" type="checkbox" checked={hold} onChange={(e)=>setHold(e.target.checked)}/><label htmlFor="hold-invoice">Hold customer reminders for this invoice</label></div><Field label="Manual escalation"><select value={escalation} onChange={(e)=>setEscalation(e.target.value)}><option value="">Use automatic escalation</option><option>CFO Escalation</option><option>Management Escalation</option><option>Legal Action</option></select></Field><p className="muted">Automatic stage: {invoice.stage}. Manual decisions are stored separately in Followup State; the formula column is preserved.</p><Field label="Follow-up note"><textarea rows={3} value={note} onChange={(e)=>setNote(e.target.value)}/></Field><Field label="Sales / RSE remarks"><textarea rows={3} value={remarks} onChange={(e)=>setRemarks(e.target.value)}/></Field><button type="submit" className="desk-btn primary">Save follow-up</button></form>
}
function ConfirmationForm({party,state,commit,setModal}:FormContext&{party:Party}) {
  const [date,setDate]=useState(today()),[status,setStatus]=useState<Confirmation["status"]>("Confirmed"),[reference,setReference]=useState(""),[note,setNote]=useState("")
  return <form className="desk-form" onSubmit={(e)=>{e.preventDefault();if(!validDate(date))return toast.error("Choose a valid date");const confirmation:Confirmation={id:crypto.randomUUID(),at:new Date().toISOString(),party:party.name,date,amount:party.total,status,reference,note};commit({...state,revision:state.revision+1,confirmations:[...state.confirmations,confirmation],activity:[...state.activity,{...addActivity(state,"Balance response recorded",null,`${status} · ${money(party.total)} · ${reference||note}`),party:party.name,after:JSON.stringify(confirmation)}]});setModal(null);toast.success("Balance response recorded")} }><h3>{party.name}</h3><div className="desk-notice">Current balance snapshot: {money(party.total)}. Recording a response does not post a payment or change the ledger.</div><div className="form-two"><Field label="Response date"><input required type="date" value={date} onChange={(e)=>setDate(e.target.value)}/></Field><Field label="Customer response"><select value={status} onChange={(e)=>setStatus(e.target.value as Confirmation["status"])}><option>Confirmed</option><option>Disputed</option></select></Field></div><Field label="Reference / received via"><input value={reference} onChange={(e)=>setReference(e.target.value)} placeholder="Email, call or document reference"/></Field><Field label="Response notes"><textarea required={status==="Disputed"} rows={4} value={note} onChange={(e)=>setNote(e.target.value)}/></Field><button type="submit" className="desk-btn primary">Record response</button></form>
}
function SettingsForm({state,commit,setModal}:FormContext) {
  const [settings,setSettings]=useState({...state.settings}),[type,setType]=useState("Reminder")
  const change=(key:keyof Settings,value:string|number)=>setSettings((s)=>({...s,[key]:value})),subject=type==="Reminder"?"reminderSubject":"balanceSubject",body=type==="Reminder"?"reminderBody":"balanceBody"
  return <form className="desk-form" onSubmit={(e)=>{e.preventDefault();if(!settings.reminderSubject.trim()||!settings.reminderBody.trim()||!settings.balanceSubject.trim()||!settings.balanceBody.trim())return toast.error("Both message types need a subject and body");commit({...state,revision:state.revision+1,settings,activity:[...state.activity,addActivity(state,"Templates / rules updated",null,"Updated message templates and reminder cadence",JSON.stringify(state.settings),JSON.stringify(settings))]});setModal(null);toast.success("Templates and rules saved")}}><div className="form-two"><Field label="Include invoices due within (days)"><input type="number" min="0" max="365" required value={settings.preDueDays} onChange={(e)=>change("preDueDays",Number(e.target.value))}/></Field><Field label="Repeat reminders after (days)"><input type="number" min="1" max="365" required value={settings.repeatDays} onChange={(e)=>change("repeatDays",Number(e.target.value))}/></Field></div><div className="segmented">{["Reminder","Balance confirmation"].map((v)=><button key={v} type="button" onClick={()=>setType(v)} className={type===v?"active":""}>{v}</button>)}</div><Field label="Subject"><input value={settings[subject]} onChange={(e)=>change(subject,e.target.value)}/></Field><Field label="Message"><textarea rows={12} value={settings[body]} onChange={(e)=>change(body,e.target.value)}/></Field><p className="template-fields">Fields: {"{{customer_name}}, {{total_outstanding}}, {{as_of_date}}, {{invoice_count}}, {{invoice_list}}, {{stage}}"}</p><p className="muted">Changes apply to new drafts. Existing drafts retain their message and recipient snapshot.</p><button type="submit" className="desk-btn primary">Save templates & rules</button></form>
}
function MessagePreview({message,onContact,note="Email sending is not connected. This is a message preview."}:{message:{email:string;subject:string;body:string};onContact?:()=>void;note?:string}) {return <div className="message-preview"><div><span>To</span><b>{message.email||"Contact needs review"}</b>{onContact&&<button className="text-action" onClick={onContact}>Update contact</button>}</div><div><span>Subject</span><b>{message.subject}</b></div><pre>{message.body}</pre><div className="desk-notice"><Mail size={16}/>{note}</div></div>}
function sendingNote(mail:EmailStatus):string {
  if(!mail.configured)return mail.problem?`Email settings need attention: ${mail.problem}`:"Email sending is not set up yet. Messages can be prepared and reviewed."
  const copy=`${mail.copyTo?.length&&mail.mode==="live"?` A copy of each email goes to ${mail.copyTo.join(", ")}.`:""}${mail.sentCopy?" A copy of each email is saved in the Sent folder.":""}`
  return mail.mode==="live"?`Live: Send delivers prepared drafts to customers from ${mail.sender}.${copy}`:`Test mode: Send delivers only to ${mail.testRecipient}. Customers receive nothing and drafts stay Prepared.${copy}`
}
function EmailBadge({status,senderName,onChangeName}:{status:EmailStatus;senderName:string;onChangeName:()=>void}) {
  if(!status.configured)return <span className={`email-status ${status.problem?"problem":"pending"}`} title={status.problem??"Mailbox settings have not been added to this environment"}><Mail size={14}/>{status.problem?"Email settings need attention":"Email setup pending"}</span>
  return <span className={`email-status ${status.mode}`} title={`Sending from ${status.sender}`}><Mail size={14}/><button className="sender-name" title="Change the name recorded with emails you send" onClick={onChangeName}>{senderName||"Add your name"}</button> · {status.mode==="live"?<>Live from {status.sender}</>:<>Test mode → {status.testRecipient}</>}</span>
}
function SendButton({drafts,mail,sending,onSend,single=false}:{drafts:Reminder[];mail:EmailStatus;sending:boolean;onSend:(drafts:Reminder[])=>void;single?:boolean}) {
  const live=mail.mode==="live"
  const label=sending?"Sending…":single?(live?"Send email":"Send test"):`${live?"Send":"Test send"} ${drafts.length} prepared`
  return <button className={`desk-btn${live?" send-live":""}`} disabled={!mail.configured||sending||!drafts.length} title={mail.configured?undefined:"Email sending is not set up yet"} onClick={()=>onSend(drafts)}><Send size={16}/>{label}</button>
}
function driveLine(sync:{state:string;message?:string;at?:string;remote?:DriveFile},dirty:boolean):string {
  if(sync.state==="saving")return "Saving to OneDrive…"
  if(sync.state==="conflict")return `The tracker was changed in OneDrive${sync.remote?.modifiedBy?` by ${sync.remote.modifiedBy}`:""} while this computer had unsaved changes. Choose which version to keep.`
  if(sync.state==="locked")return sync.message??"The tracker is open for editing elsewhere. Saving will retry."
  if(sync.state==="error")return `Not saved to OneDrive: ${sync.message}`
  if(dirty)return "Saving to OneDrive in a moment…"
  return `Saved to OneDrive${sync.at?` · ${new Date(sync.at).toLocaleString("en-IN",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"})}`:""}`
}
function OneDrivePanel({drive,busy,hasWorkbook,onUse,onCreate,onDisconnect}:{drive:OneDriveStatus;busy:boolean;hasWorkbook:boolean;onUse:(file:DriveFile)=>void;onCreate:()=>void;onDisconnect:()=>void}) {
  const [files,setFiles]=useState<DriveFile[]|null>(null),[error,setError]=useState("")
  useEffect(()=>{if(drive.connected)listDriveFiles().then(setFiles).catch((e:unknown)=>{setFiles([]);setError(e instanceof Error?e.message:"OneDrive could not be reached")})},[drive.connected])
  if(!drive.configured)return <div className="onedrive-panel"><p>OneDrive is not set up in this environment yet. Add the OneDrive settings and restart the app.</p></div>
  if(!drive.connected)return <div className="onedrive-panel"><p>Sign in once with the Microsoft account that will hold the shared tracker. After that, nobody else needs to sign in.</p><p className="muted">The app can only use its own folder in that OneDrive: <b>Apps › JAL Receivables Desk</b>. It cannot see any other files.</p>{drive.error&&<div className="desk-notice">{drive.error}</div>}<a className="desk-btn primary" href="/api/onedrive/connect"><Cloud size={16}/>Connect OneDrive</a></div>
  return <div className="onedrive-panel">
    <p className="muted">Connected as <b>{drive.account}</b>. Trackers live in <b>Apps › JAL Receivables Desk</b> in that OneDrive.</p>
    {error&&<div className="desk-notice">{error}</div>}
    {files===null?<p className="muted">Loading files…</p>:files.length?<ul className="onedrive-files">{files.map((f)=><li key={f.id}><div><b>{f.name}</b><small>Changed {new Date(f.modified).toLocaleString("en-IN",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"})}{f.modifiedBy?` by ${f.modifiedBy}`:""}</small></div>{drive.file?.id===f.id?<Status value="In use"/>:<button className="desk-btn" disabled={busy} onClick={()=>onUse(f)}>Use this tracker</button>}</li>)}</ul>:<p className="muted">No Excel files in the folder yet.{hasWorkbook?" Put the open workbook there to start sharing it.":" Open your tracker from this computer first, then put it into OneDrive from here."}</p>}
    <div className="onedrive-actions">{hasWorkbook&&<button className="desk-btn primary" disabled={busy} onClick={onCreate}><Upload size={16}/>Put the open workbook into OneDrive</button>}<button className="text-action" onClick={onDisconnect}>Disconnect OneDrive</button></div>
  </div>
}
function Field({label,children}:{label:string;children:React.ReactNode}) {return <label className="desk-field"><span>{label}</span>{children}</label>}

function JalLogo() { return <Image className="jal-logo" src="/jal-logo.png" alt="JAL" width={266} height={108} unoptimized priority/> }
