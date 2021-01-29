# Import the SwisPowerShell module
Import-Module -Name SwisPowerShell

# Name the Swis connection variable $swis

#$swis = Connect-Swis -Hostname "localhost" -Trusted

# Read in the CSV File, Adjust the path if required !!!!
# Check if the file is really a CSV and not delimited with semicolons !!!!!!!
$Nodes = import-csv -path "C:\Users\John\Desktop\Russian Printers\All_matrix_printers_RUSSIA.csv" -Delimiter ";"
# -Delimiter ";"

$N = 0


Foreach ($Node in $Nodes)

{
Write-host ($Node.IP)
# Check if the IP Address is already in Solarwinds
        #$NID = Get-SwisData $swis "SELECT NodeID FROM Orion.Nodes WHERE IPAddress = '$($Node.IP)'"
        # Only create a node if its IP Address does not exist already
        If (TRUE)
        {
        $N += 1
# Regions WER and CEGB run on the Main Poller, all other regions on the APE
        If (($Node.Region -eq 'WER') -or ($Node.Region -eq 'CEGB'))
            {$EngineID = 5}
        else
            {$EngineID = 6}    
# Set the Node properties    
        $properties = @{
            "EngineID" = $EngineID
            "Caption" = $Node."IP"

            "IPAddress" = $Node."IP"

# ObjectSubType ICMP

       	    "ObjectSubType" = "SNMP"
    	    "SNMPVersion" = 2
    	    "Community" = "public"






        }

# Create the new Node
        $newURI = New-SwisObject -Swisconnection $swis -EntityType "Orion.Nodes" -Properties $properties


# Define SNMP Pollers 
        $pollers = @(
        "N.Details.SNMP.Generic",
        "N.Uptime.SNMP.Generic",
        "N.Status.ICMP.Native",
        "N.Status.SNMP.Native",
        "N.ResponseTime.ICMP.Native",
        "N.ResponseTime.SNMP.Native",
        "N.Topology_Layer3.SNMP.ipNetToMedia",
        "I.Rediscovery.SNMP.IfTable",
        "I.StatisticsErrors32.SNMP.IfTable",
        "I.StatisticsTraffic.SNMP.Universal",
        "I.Status.SNMP.IfTable"
        )

# Get the NodeID of the newly added Node
        $nodeID = $newURI.split('=')[1]

# Set the Poller properties
        foreach ($poller in $pollers) {
                $properties = @{
                "PollerType" = $poller
                "NetObject" = "N:" + $NodeID
                "NetObjectType" = "N"
                "NetObjectID" = $nodeID
                "Enabled" = "True" 
        }

# Add the ICMP Pollers to the new Node
        $newURI = New-SwisObject -SwisConnection $swis -EntityType "Orion.Pollers" -Properties $properties

    }
#    Update the Customer Properties City, Country, Region, Department, Importance, Type_of_Device, Comments, Owner and BIA

# Retrieve the NodeID and URI of the new Node
    $row = Get-SwisData $swis "SELECT NodeID, URI FROM Orion.Nodes WHERE Caption = '$($Node."IP")'"

# Set the Custom property values

            $City = $Node."City"
            $Country = $Node."Country"
            $Region = $Node."Region"
            $Department = $Node."Department"
            $Importance = $Node."Importance"
            $Type_of_device = $Node."Type_of_Device"
            $SiteID = $Node."Site ID"
            $SerialNumber = $Node."Serial Number"
            $BusinessImpactedArea = $Node."Business Impacted Area"
            $Owner = $Node."Owner"
        
           



# Uncomment if you want to display the retrieved values on the Console
            Write-Host "City: $City"
            Write-Host "Country: $Country"
            Write-Host "Region: $Region"
            Write-Host "Type_Of_Device: $Type_Of_Device"
            Write-host "Department : $Department"
            Write-host "Importance : $Importance"
            Write-host "Type_of_device = $Type_of_Device"
            Write-host "SiteID = $SiteID"
            Write-host "Owner : $Owner"
            Write-host "BIA : $BusinessImpactedArea"

# Set the custom Properties

            $Properties = @{
                 "City" = $City
                 "Country" = $Country
                 "Region" = $Region
                 "Department" = $Department
                 "Type_Of_Device" = $Type_Of_Device
                 "Importance" = $Importance
                 "SiteID" = $SiteID
                 "SerialNumber" = $Node."SerialNumber"
                 "Owner" = $Owner
                 "BusinessImpactedArea" = $BusinessImpactedArea
                }

# Set the Custom Properties URI

            $CustomPropertiesURI = $row.URI + "/CustomProperties"

# Add the Custom Properties to the New Node

            Set-SwisObject -SwisConnection $swis -Uri $CustomPropertiesURI  -Properties $Properties

# Display NodeName and IP Address of the New Node on the Console

    Write-Host "Node $($Node."DisplayName" ) with IP Address $($Node."IP ") has been created"
    break
   }
break
}
# Finally write the total number of nodes created on the Console

Write-Host "Total number of Nodes created: $($N)"

