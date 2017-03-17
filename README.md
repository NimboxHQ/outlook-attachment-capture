# What is this?

This tool for Outlook desktop will capture inbound email attachments, and save them to the user's Vault.

# Client-side setup instructions

1. Open the VBA IDE (Alt-F11) in the user’s Outlook.
2. Insert the `nimbox.vbs` code to the Modules section.
  * On the left there is a tree, expand until you find *Modules*. Then, if there is not a *Module* item under Modules, create one by right clicking on Modules, and selecting **Insert** > **Module**.
3. On line 6 of `nimbox.vbs` edit `c:\[user]\Nimbox Vault\Outlook Attachments` to reflect the user’s profile folder path.
4. Close the VBA IDE.
5. Create a Rule that calls the script: **Tools** > **Rules and Alerts** > **New Rule**.
  * In the first screen of the new rule wizard, choose *Check messages when they arrive*.
  * In the second, you may specify certain criteria that the message must match.
  * In the third screen, choose *Run a Script*. When you click the underlined word, *script*, you should see the code that you pasted in the VBA console.
6. Click **Finish**, and test.
