# Remove existing NAT network
Get-NetNat | Remove-NetNat -Confirm:$false

# Create a new Internal Virtual Switch
New-VMSwitch -Name "NATSwitch" -SwitchType Internal

# Set IP address for the new switch
New-NetIPAddress -IPAddress 192.168.1.1 -PrefixLength 24 -InterfaceAlias "vEthernet (NATSwitch)"

# Create a new NAT network
New-NetNat -Name "VMNat" -InternalIPInterfaceAddressPrefix 192.168.1.0/24
