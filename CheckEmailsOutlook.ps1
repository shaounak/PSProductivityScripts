$message  = 'Outlook Item'
$question = 'What would you like to do?'
$choices  = 'Mark as &Read', '&Flag Item', '&Mark as Read and Flag Item', '&Next Item', '&Close'

$outlook = new-object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
$folder=$namespace.GetDefaultFolder(6)
$folder.Items | 
	?{$_.Unread -eq $true } |
	sort receivedtime | 
	%{
		Clear-Host
        $item = $_
        $message="=============================================================================================================================================================="
        $message += "`nSenderName : $($_.SenderName)"
        $message += "`nConversationTopic : $($_.ConversationTopic)"
        $message += "`nReceivedTime : $($_.ReceivedTime)"
        $message += "`nTo : $($_.To)"
        $message += "`nCC : $($_.CC)"
        $message += "`n=============================================================================================================================================================="
        $message += $_.body 
        $message | more
        Write-Host "=============================================================================================================================================================="
        $decision = $Host.UI.PromptForChoice($message, $question, $choices, 4)
        switch ($decision) {
        0 {	$item.Unread = $false; $item.close(0); }
        1 { $item.flagStatus = 2; $item.FlagIcon = 6; $item.close(0); }
        2 { $item.Unread = $false; $item.flagStatus = 2; $item.FlagIcon = 6; $item.close(0); }
        3 { Write-Host "Fetching Next Item..."; }
        4 { Write-Host "Closing..."; break;}
        }
        if($decision -eq 4){break;}
	}
Write-Host -NoNewline 'Press any key to continue...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
