# Mirror-AD-Groups-to-Entra-Groups

This script was written for mirroring groups from an on-prem enviroment to Entra/Microsofot 365. The users who were members of the groups were not synced with Entra Connect and needed to be looked up through email address.

This script can be used to add members of on-prem AD groups to Entra ID groups based off of the members email address. It will prompt the user to select an excel .xlsx file containing the groups with the below headers:

  ADGroupName | EntraGroupName
 -------------+----------------
  Group       | Group 
