# DragDropDemo v2.1

(Project update: V2.1 updates to use Coclass syntax, requires twinBASIC Beta 167 or newer)
(Project update: V2 fixes incorrect IEnumSTATDATA definition, and since LongLong can be used on both x86 and x64, ditches the separate definitions in favor of using LongLong+CopyMemory on both)

A while back I posted this project:

[[VB6] Register any control as a drop target that shows the Explorer drag image](https://www.vbforums.com/showthread.php?808125-VB6-Register-any-control-as-a-drop-target-that-shows-the-Explorer-drag-image)

![img](http://i.imgur.com/y3SHMsH.jpg) ![img2](http://i.imgur.com/aUaniDK.jpg)


I wanted to try my hand at using these interfaces in x64 apps. A 64-bit version of oleexp doesn't seem likely; there's seemingly insurmountable barriers to using midl with a project that needs to redefine things in the force-included headers, and the original types weren't preserved so every single interface would have to be manually reviewed and updated if it uses pointer types. But, twinBASIC lets you define interfaces as a native language feature. So I set about re-implementing all the necessary ones for this project:

```
[ InterfaceId ("4657278B-411B-11D2-839A-00C04FD918D0") ]
Interface IDropTargetHelper Extends stdole.IUnknown
    Sub DragEnter(ByVal hwndTarget As LongPtr, ByVal pDataObject As DragDropDemo.IDataObject, ppt As POINT, ByVal dwEffect As DROPEFFECTS)
    Sub DragLeave()
    Sub DragOver(ppt As POINT, ByVal dwEffect As DROPEFFECTS)
    Sub Drop(ByVal pDataObject As DragDropDemo.IDataObject, ppt As POINT, ByVal dwEffect As DROPEFFECTS)
    Sub Show(ByVal fShow As Long)
End Interface

[ InterfaceId ("DE5BF786-477A-11D2-839D-00C04FD918D0") ]
Interface IDragSourceHelper Extends stdole.IUnknown
    Sub InitializeFromBitmap(pshdi As SHDRAGIMAGE, pDataObject As DragDropDemo.IDataObject)
    Sub InitializeFromWindow(ByVal hwnd As LongPtr, ppt As POINT, pDataObject As DragDropDemo.IDataObject)
End Interface

[ InterfaceId ("83E07D0D-0C5F-4163-BF1A-60B274051E40") ]
Interface IDragSourceHelper2 Extends IDragSourceHelper
	Sub SetFlags(ByVal dwFlags As DSH_FLAGS)
End Interface

[ InterfaceId ("00000122-0000-0000-C000-000000000046") ]
#If Win64 Then
Interface IDropTarget Extends stdole.IUnknown
    Sub DragEnter(ByVal pDataObject As DragDropDemo.IDataObject, ByVal grfKeyState As Long, ByVal pt As LongPtr, dwEffect As DROPEFFECTS)
    Sub DragOver(ByVal grfKeyState As Long, ByVal pt As LongPtr, pdwEffect As DROPEFFECTS)
    Sub DragLeave()
    Sub Drop(ByVal pDataObj As DragDropDemo.IDataObject, ByVal grfKeyState As Long, ByVal pt As LongPtr, pdwEffect As DROPEFFECTS)
End Interface
#Else
Interface IDropTarget Extends stdole.IUnknown
    Sub DragEnter(ByVal pDataObject As DragDropDemo.IDataObject, ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY as long, dwEffect As DROPEFFECTS)
    Sub DragOver(ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY as long, pdwEffect As DROPEFFECTS)
    Sub DragLeave()
    Sub Drop(ByVal pDataObj As DragDropDemo.IDataObject, ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY as long, pdwEffect As DROPEFFECTS)
End Interface
#End If

[ InterfaceId ("00000121-0000-0000-C000-000000000046") ]
Interface IDropSource Extends stdole.IUnknown
    Sub QueryContinueDrag(ByVal fEscape As Long)
    Sub GiveFeedback(ByVal grfKeyState As Long)
End Interface

[ InterfaceId ("0000010f-0000-0000-C000-000000000046") ]
Interface IAdviseSink Extends stdole.IUnknown
	Sub OnDataChange(pFormatEtc As FORMATETC, pStgMed As STGMEDIUM)
    Sub OnViewChange(dwAspect As DVASPECT, ByVal lindex As Long)
    Sub OnRename(ByVal pmk As LongPtr) 'As IMoniker
    Sub OnSave()
    Sub OnClose()
End Interface

[ InterfaceId ("00000103-0000-0000-C000-000000000046") ]
Interface IEnumFormatETC Extends stdole.IUnknown
    Sub Next(ByVal celt As Long, rgelt As FORMATETC, pceltFetched As Long)
    Sub Skip(ByVal celt As Long)
    Sub Reset()
    Sub Clone(ppEnum As DragDropDemo.IEnumFormatETC)
End Interface

[ InterfaceId ("00000105-0000-0000-C000-000000000046") ]
Interface IEnumSTATDATA Extends stdole.IUnknown
    Sub Next(ByVal celt As Long, rgelt As STATDATA, pceltFetched As Long)
    Sub Skip(ByVal celt As Long)
    Sub Reset()
    Sub Clone(ppEnum As DragDropDemo.IEnumSTATDATA)
End Interface

[ InterfaceId ("0000010E-0000-0000-C000-000000000046") ]
Interface IDataObject Extends stdole.IUnknown
    Sub GetData(pFormatEtcIn As FORMATETC, pMedium As STGMEDIUM)
    Sub GetDataHere(pFormatEtc As FORMATETC, pMedium As STGMEDIUM)
    Sub QueryGetData(pFormatEtc As FORMATETC)
    Sub GetCanonicalFormatEtc(pFormatEtcIn As FORMATETC, pFormatEtcOut As FORMATETC)
    Sub SetData(pFormatEtc As FORMATETC, pMedium As STGMEDIUM, ByVal fRelease As Long)
    Function EnumFormatEtc(ByVal dwDirection As DATADIR) As DragDropDemo.IEnumFormatETC
    Sub DAdvise(pFormatEtc As FORMATETC, ByVal advf As ADVF, pAdvSink As DragDropDemo.IAdviseSink)
    Sub DUnadvise(ByVal dwConnection As Long)
    Function EnumDAdvise() As DragDropDemo.IEnumSTATDATA
End Interface

'CLSID_DragDropHelper 
[ CoClassId ("4657278A-411B-11D2-839A-00C04FD918D0") ]
[ COMCreatable ]
CoClass DragDropHelper
	 [ Default ] Interface IDropTargetHelper
	Interface IDragSourceHelper
	Interface IDragSourceHelper2
End CoClass
```

I used functions for some to preserve as much compatibility as possible with oleexp-using code.

There were a couple tricky things here... it's always been an odd interface. Normally, UDTs as [in] are ByRef; MSDN lists the x,y coords as POINT, and we had to expand that to 2 ByVal Long's in x86. It's even weirder in x64; we have to use a single member, but it's also ByVal... so it's 8 bytes but not *actually* a pointer, so it's handled like this:

```
#If Win64 Then
Private Sub IDropTarget_DragEnter(ByVal pDataObj As DragDropDemo.IDataObject, ByVal grfKeyState As Long, ByVal pt As LongPtr, pdwEffect As DragDropDemo.DROPEFFECTS)
    Dim ptt As DragDropDemo.POINT
    CopyMemory ptt, pt, LenB(ptt)
#Else
Private Sub IDropTarget_DragEnter(ByVal pDataObj As DragDropDemo.IDataObject, ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY As Long, pdwEffect As DragDropDemo.DROPEFFECTS)
    Dim ptt As DragDropDemo.POINT
    ptt.x = ptX: ptt.y = ptY
#End If
```

Also, lots of things are fully qualified as there's currently an issue with conflicting types in the WinNativeForms package that are exposed to users.

But besides those quirks, the code works on x64 without major change, just updating to LongPtr where needed, and replacing the PictureBox with a Frame until that's available.

Et voilà:

![Imgur](https://i.imgur.com/gysxo6r.jpg)

### Requirements
Windows Vista or newer
[twinBASIC Beta 167 or newer](https://github.com/twinbasic/twinbasic/releases)

Thanks to twinBASIC developer Wayne Phillips for his help getting this working, and of course for the continuing great work on twinBASIC itself :thumb:
