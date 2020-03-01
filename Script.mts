DataTable.ImportSheet "C:\Users\NHIDCL\Documents\FlightLogin.xlsx", 1, "Global"


username = DataTable("Username", dtGlobalSheet)'"john"
varPassword = DataTable("Password", dtGlobalSheet)'"HP"
EncryptedPWD = Crypt.Encrypt(varPassword)


LoadFunctionLibrary "C:\Users\NHIDCL\Documents\Unified Functional Testing\Library1.qfl"

Call Login(username,EncryptedPWD)


