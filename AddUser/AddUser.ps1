#Nhập module hoạt động
Import-Module ActiveDirectory
$file = Read-Host -Prompt "Nhập đường dẫn file CSV"
#Chèn file danh sách User
$csv = import-csv $file
foreach ($row in $csv)
{
    $name = $row.firstname
    $ho = $row.lastname
    $path = $row.ou
    $account = $row.account
    $fullname = $row.displayname
    $passrord = $row.password
    $username = $row.account+"@cnttcoitk14.vn"
#Kiểm tra xem các User trong file đã có chưa
if (Get-ADUser -F {SamAccountName -eq $account})
#Nếu có rồi thì in ra thông báo màu vàng
{Write-Host -ForegroundColor Yellow "Tài khoản $account đã có người dùng !"}
#Còn nếu chưa có thì bắt đầu thêm vào
else
{
    New-ADUser -DisplayName "$fullname" `
    -Name "$fullname" `
    -GivenName "$name" `
    -Surname "$ho" `
    -UserPrincipalName "$username" `
    -SamAccountName "$account" `
    -AccountPassword (ConvertTo-SecureString "$passrord" -AsPlainText -Force) `
    -ChangePasswordAtLogon $false `
    -Enabled $true `
    -Path $path `
    #Thêm thành công thì in ra thông báo màu xanh
    Write-Host -ForegroundColor Green “Da Them thanh cong $ho $name”
}
}