#Nhập module hoạt động
Import-Module ActiveDirectory
$file = Read-Host -Prompt "Nhap duong dan file CSV"
#Chèn file danh sách User
$csv = import-csv $file
foreach ($row in $csv)
{
    $account = $row.account
    $password = $row.password
#Kiểm tra xem các User trong file đã có chưa
if (!(Get-ADUser -F {SamAccountName -eq $account}))
#Nếu có rồi thì in ra thông báo màu vàng
{Write-Host -ForegroundColor Yellow "Tai khoan $account khong ton tai !"}
#Nếu tồn tại, thay đổi mật khẩu
else
{
    Set-ADAccountPassword -Identity "$account" `
    -NewPassword (ConvertTo-SecureString -AsPlainText "$password" -Force) -Reset `
    #Thay thành công thì in ra thông báo màu xanh
    Write-Host -ForegroundColor Green “Da Thay Doi PassWord Thanh Cong $account”
}
}