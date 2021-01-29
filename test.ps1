# Read in the CSV File, Adjust the path if required !!!!
# Check if the file is really a CSV and not delimited with semicolons !!!!!!!
$Nodes = import-csv -path "C:\Users\John\Desktop\Russian Printers\All_matrix_printers_RUSSIA.csv"
# -Delimiter ";"
# Set the Node properties   
Foreach ($Node in $Nodes) {
    If (($Node.IP -ne 'No') -and ($Node.IP -ne 'USB')) {
        $properties = @{
            "EngineID"      = 6
            "Caption"       = $Node.IP
            "IPAddress"     = $Node."IP"
            # ObjectSubType ICMP
            "ObjectSubType" = "SNMP"
            "SNMPVersion"   = 2
            "Community"     = "public"
        }
        Write-Host($Node.IP)
    }
}
