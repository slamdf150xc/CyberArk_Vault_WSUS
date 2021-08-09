# CyberArk_Vault_WSUS
Script to configure CyberArk Vault to Download and Install updates from a WSUS server.

## Script Usage
```
.\CyberArk_WSUS.ps1
```

You will be prompted with a menu...
```
================ WSUS Options ================

Your WSUS server is currently set to: http://10.1.20.12:8530

1: Configure WSUS server URL:Port
2: Start WSUS services and open the firewall
3: Stop WSUS servies and close the firewall
4: Download updates from WSUS
5: Install updates that have been downloaded
6: Download then Install updates from WSUS
7: Reboot Server
8: Force Vault to check in with WSUS
Q: Press 'Q' to quit.

Please make a selection:
```

If you've not setup WSUS on your Vault server you will be prompted to configure the WSUS URL.
```
It looks like you've not setup your WSUS URL. Would you like to do that now?
(Y/N):
```

If you choose "No" you will need to select Option 1 when the menu loads.