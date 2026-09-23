Attribute VB_Name = "modCommands"
Option Explicit

Sub DangNhap()
    frmDangNhap.Show
End Sub

Sub moUFTrichXuat()
    frmTrichXuatXML.Show
End Sub

Sub closeForm()
    Unload frmDangNhap
    frmTaiHoaDon.Show
End Sub

Sub TaiHoaDon()
    frmTaiHoaDon.Show
End Sub
