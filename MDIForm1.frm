VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10470
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   17490
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu swd 
      Caption         =   "Patient"
      Begin VB.Menu fgbfg 
         Caption         =   "New Patient"
      End
      Begin VB.Menu tyhyth 
         Caption         =   "Update Patient"
      End
      Begin VB.Menu hggh 
         Caption         =   "Delete Patient"
      End
      Begin VB.Menu sdfg 
         Caption         =   "Search Patient"
      End
   End
   Begin VB.Menu dfdf 
      Caption         =   "Staff"
      Begin VB.Menu rgterg 
         Caption         =   "New Staff"
      End
      Begin VB.Menu fffff 
         Caption         =   "Update Staff"
      End
      Begin VB.Menu hjjj 
         Caption         =   "Delete Staff"
      End
      Begin VB.Menu uiyuiujk 
         Caption         =   "Search Staff"
      End
   End
   Begin VB.Menu grffg 
      Caption         =   "Doctor"
      Begin VB.Menu thhrfghg 
         Caption         =   "New Doctor"
      End
      Begin VB.Menu dfff 
         Caption         =   "Update Doctor"
      End
      Begin VB.Menu yjyjjh 
         Caption         =   "Delete Doctor"
      End
      Begin VB.Menu tht 
         Caption         =   "Search Doctor"
      End
   End
   Begin VB.Menu yhjyjhyj 
      Caption         =   "Bill"
      Begin VB.Menu hjgjg 
         Caption         =   "New Bill"
      End
      Begin VB.Menu gerg 
         Caption         =   "Update Bill"
      End
      Begin VB.Menu fthgfh 
         Caption         =   "Delete Bill"
      End
      Begin VB.Menu gfghg 
         Caption         =   "Search Bill"
      End
   End
   Begin VB.Menu erfet 
      Caption         =   "Pathalogy"
      Begin VB.Menu rtgrt 
         Caption         =   "New Pathalogy"
      End
      Begin VB.Menu hgjfg 
         Caption         =   "Update Pathalogy"
      End
      Begin VB.Menu sdfsdgf 
         Caption         =   "Delete Pathalogy"
      End
      Begin VB.Menu gbvdfg 
         Caption         =   "Search Pathalogy"
      End
   End
   Begin VB.Menu ertgr 
      Caption         =   "Covid Section"
      Begin VB.Menu rytery 
         Caption         =   "New Covid"
      End
      Begin VB.Menu tyrt 
         Caption         =   "Update Covid"
      End
      Begin VB.Menu thrtgh 
         Caption         =   "Delete Covid"
      End
      Begin VB.Menu ytujtyj 
         Caption         =   "Search Covid"
      End
   End
   Begin VB.Menu fdbht 
      Caption         =   "Reports"
      Begin VB.Menu nhtgn 
         Caption         =   "Patient Report"
      End
      Begin VB.Menu tghrtgh 
         Caption         =   "Staff Report"
      End
      Begin VB.Menu rgerg 
         Caption         =   "Doctor Report"
      End
      Begin VB.Menu gfdb 
         Caption         =   "Bill Report"
      End
      Begin VB.Menu jtrgyj 
         Caption         =   "Pathology Report"
      End
      Begin VB.Menu thghg 
         Caption         =   "Covid Report"
      End
   End
   Begin VB.Menu DFGG 
      Caption         =   "About"
   End
   Begin VB.Menu thh 
      Caption         =   "Other"
      Begin VB.Menu fgg 
         Caption         =   "Notepad"
      End
      Begin VB.Menu dgdgg 
         Caption         =   "Calculator"
      End
   End
   Begin VB.Menu hyth 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dfff_Click()
updatedoctor.Show
End Sub

Private Sub DFGG_Click()
frmAbout.Show
End Sub

Private Sub dgdgg_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub fffff_Click()
updatestaff.Show
End Sub

Private Sub fgbfg_Click()
newpatient.Show
End Sub

Private Sub fgg_Click()
Shell "notepad.exe", vbNormalFocus
End Sub

Private Sub fthgfh_Click()
deletebill.Show
End Sub

Private Sub gbvdfg_Click()
searchpathology.Show
End Sub

Private Sub gerg_Click()
updatebill.Show
End Sub

Private Sub gfdb_Click()
DataReport4.Show
End Sub

Private Sub gfghg_Click()
searchbill.Show
End Sub

Private Sub hggh_Click()
deletepatient.Show
End Sub

Private Sub hgjfg_Click()
updatepathology.Show
End Sub

Private Sub hjgjg_Click()
addbill.Show
End Sub

Private Sub hjjj_Click()
deletestaff.Show
End Sub

Private Sub hyth_Click()
End
End Sub

Private Sub jtrgyj_Click()
DataReport5.Show
End Sub

Private Sub nhtgn_Click()
DataReport1.Show
End Sub

Private Sub rgerg_Click()
DataReport3.Show
End Sub

Private Sub rgterg_Click()
addstaff.Show
End Sub

Private Sub rtgrt_Click()
addpathology.Show
End Sub

Private Sub rytery_Click()
addcovid.Show
End Sub

Private Sub sdfg_Click()
searchpatient.Show
End Sub

Private Sub sdfsdgf_Click()
deletepathology.Show
End Sub

Private Sub tghrtgh_Click()
DataReport2.Show
End Sub

Private Sub thghg_Click()
DataReport6.Show
End Sub

Private Sub thhrfghg_Click()
adddoctor.Show
End Sub

Private Sub thrtgh_Click()
deletecovid.Show
End Sub

Private Sub tht_Click()
searchdoctor.Show
End Sub

Private Sub tyhyth_Click()
updatepatient.Show
End Sub

Private Sub tyrt_Click()
updatecovid.Show
End Sub

Private Sub uiyuiujk_Click()
ssearchstaff.Show
End Sub

Private Sub yjyjjh_Click()
deletedoctor.Show
End Sub

Private Sub ytujtyj_Click()
searchcovid.Show
End Sub
