# CyberArk_Vault_WSUS
Script to configure CyberArk Vault to Download and Install updates from a WSUS server.

## Script Usage
### Execution
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

### Menu options
* 1: Configure WSUS server URL:Port
    * You will be prompted to enter you WSUS Server URL:Port combination.
    ```
    Please enter the WSUS IP/URL and port. (http://10.1.20.12:8530):
    ```
    * If you already have a WSUS URL:Port setup you will be shown the current configuration them prompted to update it.
    ```
    Your current WSUS address is: http://10.1.20.12:8530

    Please enter the WSUS IP/URL and port. (http://10.1.20.12:8530):
    ```
* 2: Start WSUS services and open the firewall
    * This option will enable/start the nessary services for WSUS to function and open the local firewall to allow the Vault to comuticate with the configured WSUS server.
* 3: Stop WSUS servies and close the firewall
    * This option will stop/disable the WSUS services and remove the local firewall rule.
* 4: Download updates from WSUS
    * This option will first enable/start the WSUS services then open the local firewall before attempting to query the WSUS server for needed updates then download the available updates. When the process completes the WSUS services will be stopped/disabled then the local firewall rule will be removed.
* 5: Install updates that have been downloaded
    * This option will attempt to install the downloaded updates.
* 6: Download then Install updates from WSUS
    * This option will first enable/start the WSUS services then open the local firewall before attempting to query the WSUS server for needed updates then download the available updates. If any were downloaded an attempt will be made to install them. When the process completes the WSUS services will be stopped/disabled then the local firewall rule will be removed.
* 7: Reboot Server
    * This option will reboot the server.
* 8: Force Vault to check in with WSUS
    * This option will first enable/start the WSUS services then open the local firewall before attempting to check in with the WSUS server. When the process completes the WSUS services will be stopped/disabled then the local firewall rule will be removed.
* Q: Press 'Q' to quit.
    * Quit the script