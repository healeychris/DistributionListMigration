# DistributionListMigration
    .NOTES
    ===========================================================================
    Created with: 	VS
    Created on:   	28/01/2022
    Created by:   	Chris Healey
    Organization: 
    Version:        1.1	
    Filename:       Recreate-DistributionGroup.ps1
    Project path:   https://github.com/healeychris/DistributionListMigration
    Org Author:     Joe Palarchio (based on Version: 1.0) 
    Org site:       https://blogs.perficient.com/2016/05/18/office-365-script-to-move-mail-distribution-groups-to-the-cloud/
    ===========================================================================
    .DESCRIPTION
    Copies attributes of a synchronized group to a placeholder group and CSV file.  After 
    initial export of group attributes, the on-premises group can have the attribute
    "AdminDescription" set to "Group_NoSync" which will stop it from be synchronized.
    The "-Finalize" switch can then be used to write the addresses to the new group and
    convert the name.  The final group will be a cloud group with the same attributes as
    the previous but with the additional ability of being able to be "self-managed".
    Once the contents of the new group are validated, the on-premises group can be deleted.
    .NOTES

RUN ORDER - :
    	.EXAMPLE #1
        	.\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -CreatePlaceHolder       # In Exchange Online

    	.EXAMPLE #2
            .\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -Contact                 # On Prem
        	
        .EXAMPLE #3
            .\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -Finalize                # In Exchange Online
