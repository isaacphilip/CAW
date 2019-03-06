#create form
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object Windows.Forms.Form
$form.Size = New-Object Drawing.Size @(220,150)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
$form.Text = "Alkane Test"
#$form.ShowInTaskbar = $False
$form.Name = "Alkane Test"
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false

#get base64 string from here: http://www.base64-image.de

#set icon for form using Base64
$base64IconString = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAT6SURBVFhHvZcPbBNVHMe/1+tdu26l3R+3zmHt2ABF2TTqcNFgNhN0glXUKIbEaCBqjMZoTJyRIDGOhAiJiqKJI5iJGMmUsIyoaEDHIpapiWRjEGBL9rfAtnZr17W93p3v2lfWXu/2Bxc+6Uvf33u/936/+/1+xzy/0yljARAEkdb0MfEcDAYGMXF6roH+/284jp21iJIIIRajKxIsmABzgWGYeEnlugqgxZxtQJYERIUQBCkGSSZLEr+rJE5mAMvy4HkLONVJ9ZhVAFkWEY34EEEuiopWwJFth4k1ko1YJLeQJZnMI4JJYUxODWBwqBs+KQvZWRawGk9niCEmmVEAWSanDoeQ51qPypL7UHXnM6i0W+moNtHQP/D82YRTg7+j43wPjDk2sHQsiYGdqwAREc6KV/FUbT0qLLRzroRPovWXbfju9BkYzfqmpjsiywHwy+qxae01bK5grsa6R3fjMedkQkUpJRXdG2Cjy+F+6SjcebQjlZEWHO5sgy9K6mw+nKXPonaJKzGmItazCRubj8JM24r+U19F9o41tm20ngZ39wG8taKAtqaJXK5Hw6GP0Xb2V3T1d+Dc4Cmc6zuBrqwHUF2Ul6lvI4/Ov1sxTn1AqgEq6KqA6f8Ae4/vwonersRJ40yh4/DX6BwdB8vfgEUWUswWxCbPoLvtRwxJdFoq5nwsV/ZMFhX6RhgLkVePI/47B2bODCPLkdNFERz3QSRuVUEirlWMTSAYChD//hDefecnVGe8JF048mEdDtKWEgtS0TdC1gxOCRxRP8YDQ7ji68XQSC9GI154R/vQd6kPXv8gpOwarH2yBZ++fRD35NDFahiZXH+iqJnVEUnEA04GhjEWIgZjd6C68hXUVTyOpUW3osBspLNmgtzArkfQTFtqdAWQZQmx8AB8Uyyq3L/h9VX3w07H5kcXWndOC6D20LoqkIQRMM4vcaAhhveuefMEBrJLsqjRFEBxwbxjHxqf2wwb7dNDDPehq/t9fLZ3IzwTtHMeaArAyCzq1m+AibbTkEbg7dmDTxrtcG9h4N5+M7Z8vwMn/TxMmk8jECXHA6iGsjWXWKzbsVrrzoULOPZtIV785k10TNhwo6MMzsIyOOzFsLJWGHVsUiZ6V3Sv1r+CtgpcrquuM43RH7Bn2IXFeYthNZH8LvlgOYSw5V4Uai5S5ui/htoCpKUaKWQtwUrOh3AskVQyJAcQImMIiCW47cEaFMZ7M2HIFSSLGk0BDL7LEGg9DdvD2HyXGy5TMO6cxsIRWB3rsHrVbrxRVkwnqSHeUkxeVaYAmn6Awy3Y8PLPqM2mHSouXfwCHq8fMb4Q5eUvoCI388HTnMfxxhrsDxrjoUAU07fTFECSAii+vRFb657AfFMBJelOt8UwLrQvQ8NfBjDx1I12U7RVYLBi+PRWfH6sBT7aNysjX6HpSCOG09N+ghlO19MkjisHy7Qt3XxAxjj6B9rRe+UsgsaluCk/n6hGjYjA0H40t+/AIU8T/ujxoKTqNZSrJrKL1qB0/CN4RmlHCjMGI5lkuVEhCi67FAXWXGQpIZkakpIti1IIU5MD8E6MQWCyYeFkWGwrUWhKVwJv4iEH/8VFf+ZWs0ZDJSiJ4lT8k0oi9enJSoZjIBmuGSaj4hMSggkRPyLqvI/MYzkbzOp0iTCrAAuBTC1P/VmmoGmEC008F9TYXOG6CKAP8B8y8tsX2eJriwAAAABJRU5ErkJggg=="
$iconimageBytes = [Convert]::FromBase64String($base64IconString)
$ims = New-Object IO.MemoryStream($iconimageBytes, 0, $iconimageBytes.Length)
$ims.Write($iconimageBytes, 0, $iconimageBytes.Length);
$alkIcon = [System.Drawing.Image]::FromStream($ims, $true)
$Form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ims).GetHIcon())

#add an image to the form using Base64
$base64ImageString = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAT6SURBVFhHvZcPbBNVHMe/1+tdu26l3R+3zmHt2ABF2TTqcNFgNhN0glXUKIbEaCBqjMZoTJyRIDGOhAiJiqKJI5iJGMmUsIyoaEDHIpapiWRjEGBL9rfAtnZr17W93p3v2lfWXu/2Bxc+6Uvf33u/936/+/1+xzy/0yljARAEkdb0MfEcDAYGMXF6roH+/284jp21iJIIIRajKxIsmABzgWGYeEnlugqgxZxtQJYERIUQBCkGSSZLEr+rJE5mAMvy4HkLONVJ9ZhVAFkWEY34EEEuiopWwJFth4k1ko1YJLeQJZnMI4JJYUxODWBwqBs+KQvZWRawGk9niCEmmVEAWSanDoeQ51qPypL7UHXnM6i0W+moNtHQP/D82YRTg7+j43wPjDk2sHQsiYGdqwAREc6KV/FUbT0qLLRzroRPovWXbfju9BkYzfqmpjsiywHwy+qxae01bK5grsa6R3fjMedkQkUpJRXdG2Cjy+F+6SjcebQjlZEWHO5sgy9K6mw+nKXPonaJKzGmItazCRubj8JM24r+U19F9o41tm20ngZ39wG8taKAtqaJXK5Hw6GP0Xb2V3T1d+Dc4Cmc6zuBrqwHUF2Ul6lvI4/Ov1sxTn1AqgEq6KqA6f8Ae4/vwonersRJ40yh4/DX6BwdB8vfgEUWUswWxCbPoLvtRwxJdFoq5nwsV/ZMFhX6RhgLkVePI/47B2bODCPLkdNFERz3QSRuVUEirlWMTSAYChD//hDefecnVGe8JF048mEdDtKWEgtS0TdC1gxOCRxRP8YDQ7ji68XQSC9GI154R/vQd6kPXv8gpOwarH2yBZ++fRD35NDFahiZXH+iqJnVEUnEA04GhjEWIgZjd6C68hXUVTyOpUW3osBspLNmgtzArkfQTFtqdAWQZQmx8AB8Uyyq3L/h9VX3w07H5kcXWndOC6D20LoqkIQRMM4vcaAhhveuefMEBrJLsqjRFEBxwbxjHxqf2wwb7dNDDPehq/t9fLZ3IzwTtHMeaArAyCzq1m+AibbTkEbg7dmDTxrtcG9h4N5+M7Z8vwMn/TxMmk8jECXHA6iGsjWXWKzbsVrrzoULOPZtIV785k10TNhwo6MMzsIyOOzFsLJWGHVsUiZ6V3Sv1r+CtgpcrquuM43RH7Bn2IXFeYthNZH8LvlgOYSw5V4Uai5S5ui/htoCpKUaKWQtwUrOh3AskVQyJAcQImMIiCW47cEaFMZ7M2HIFSSLGk0BDL7LEGg9DdvD2HyXGy5TMO6cxsIRWB3rsHrVbrxRVkwnqSHeUkxeVaYAmn6Awy3Y8PLPqM2mHSouXfwCHq8fMb4Q5eUvoCI388HTnMfxxhrsDxrjoUAU07fTFECSAii+vRFb657AfFMBJelOt8UwLrQvQ8NfBjDx1I12U7RVYLBi+PRWfH6sBT7aNysjX6HpSCOG09N+ghlO19MkjisHy7Qt3XxAxjj6B9rRe+UsgsaluCk/n6hGjYjA0H40t+/AIU8T/ujxoKTqNZSrJrKL1qB0/CN4RmlHCjMGI5lkuVEhCi67FAXWXGQpIZkakpIti1IIU5MD8E6MQWCyYeFkWGwrUWhKVwJv4iEH/8VFf+ZWs0ZDJSiJ4lT8k0oi9enJSoZjIBmuGSaj4hMSggkRPyLqvI/MYzkbzOp0iTCrAAuBTC1P/VmmoGmEC008F9TYXOG6CKAP8B8y8tsX2eJriwAAAABJRU5ErkJggg=="
$imageBytes = [Convert]::FromBase64String($base64ImageString)
$ms = New-Object IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$ms.Write($imageBytes, 0, $imageBytes.Length);
$alkanelogo = [System.Drawing.Image]::FromStream($ms, $true)

$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width =  $alkanelogo.Size.Width;
$pictureBox.Height =  $alkanelogo.Size.Height; 
$pictureBox.Location = New-Object System.Drawing.Size(85,20) 
$pictureBox.Image = $alkanelogo;
$form.Controls.Add($pictureBox)

$btn = New-Object System.Windows.Forms.Button
$btn.AutoSize = $True
$btn.Text = "Send Email"
$btn.Location = New-Object System.Drawing.Size(65,60)
$btn.add_click(
{

	Try
	{
		#SMTP server name
		$smtpServer = "xxx.yyy.zzz";

		#Creating a Mail object
		$msg = new-object Net.Mail.MailMessage;

		#Creating SMTP server object
		$smtp = new-object Net.Mail.SmtpClient($smtpServer);

		#create a content type object for the attachment
		$ct = new-object Net.Mime.ContentType

		#create Alkane icon in memory stream
		$AlkaneBase64String = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAT6SURBVFhHvZcPbBNVHMe/1+tdu26l3R+3zmHt2ABF2TTqcNFgNhN0glXUKIbEaCBqjMZoTJyRIDGOhAiJiqKJI5iJGMmUsIyoaEDHIpapiWRjEGBL9rfAtnZr17W93p3v2lfWXu/2Bxc+6Uvf33u/936/+/1+xzy/0yljARAEkdb0MfEcDAYGMXF6roH+/284jp21iJIIIRajKxIsmABzgWGYeEnlugqgxZxtQJYERIUQBCkGSSZLEr+rJE5mAMvy4HkLONVJ9ZhVAFkWEY34EEEuiopWwJFth4k1ko1YJLeQJZnMI4JJYUxODWBwqBs+KQvZWRawGk9niCEmmVEAWSanDoeQ51qPypL7UHXnM6i0W+moNtHQP/D82YRTg7+j43wPjDk2sHQsiYGdqwAREc6KV/FUbT0qLLRzroRPovWXbfju9BkYzfqmpjsiywHwy+qxae01bK5grsa6R3fjMedkQkUpJRXdG2Cjy+F+6SjcebQjlZEWHO5sgy9K6mw+nKXPonaJKzGmItazCRubj8JM24r+U19F9o41tm20ngZ39wG8taKAtqaJXK5Hw6GP0Xb2V3T1d+Dc4Cmc6zuBrqwHUF2Ul6lvI4/Ov1sxTn1AqgEq6KqA6f8Ae4/vwonersRJ40yh4/DX6BwdB8vfgEUWUswWxCbPoLvtRwxJdFoq5nwsV/ZMFhX6RhgLkVePI/47B2bODCPLkdNFERz3QSRuVUEirlWMTSAYChD//hDefecnVGe8JF048mEdDtKWEgtS0TdC1gxOCRxRP8YDQ7ji68XQSC9GI154R/vQd6kPXv8gpOwarH2yBZ++fRD35NDFahiZXH+iqJnVEUnEA04GhjEWIgZjd6C68hXUVTyOpUW3osBspLNmgtzArkfQTFtqdAWQZQmx8AB8Uyyq3L/h9VX3w07H5kcXWndOC6D20LoqkIQRMM4vcaAhhveuefMEBrJLsqjRFEBxwbxjHxqf2wwb7dNDDPehq/t9fLZ3IzwTtHMeaArAyCzq1m+AibbTkEbg7dmDTxrtcG9h4N5+M7Z8vwMn/TxMmk8jECXHA6iGsjWXWKzbsVrrzoULOPZtIV785k10TNhwo6MMzsIyOOzFsLJWGHVsUiZ6V3Sv1r+CtgpcrquuM43RH7Bn2IXFeYthNZH8LvlgOYSw5V4Uai5S5ui/htoCpKUaKWQtwUrOh3AskVQyJAcQImMIiCW47cEaFMZ7M2HIFSSLGk0BDL7LEGg9DdvD2HyXGy5TMO6cxsIRWB3rsHrVbrxRVkwnqSHeUkxeVaYAmn6Awy3Y8PLPqM2mHSouXfwCHq8fMb4Q5eUvoCI388HTnMfxxhrsDxrjoUAU07fTFECSAii+vRFb657AfFMBJelOt8UwLrQvQ8NfBjDx1I12U7RVYLBi+PRWfH6sBT7aNysjX6HpSCOG09N+ghlO19MkjisHy7Qt3XxAxjj6B9rRe+UsgsaluCk/n6hGjYjA0H40t+/AIU8T/ujxoKTqNZSrJrKL1qB0/CN4RmlHCjMGI5lkuVEhCi67FAXWXGQpIZkakpIti1IIU5MD8E6MQWCyYeFkWGwrUWhKVwJv4iEH/8VFf+ZWs0ZDJSiJ4lT8k0oi9enJSoZjIBmuGSaj4hMSggkRPyLqvI/MYzkbzOp0iTCrAAuBTC1P/VmmoGmEC008F9TYXOG6CKAP8B8y8tsX2eJriwAAAABJRU5ErkJggg=="
		$AlkaneImageBytes = [Convert]::FromBase64String($AlkaneBase64String)
		$AlkaneMs = New-Object IO.MemoryStream($AlkaneImageBytes, 0, $AlkaneImageBytes.Length)
		$AlkaneMs.Write($AlkaneImageBytes, 0, $AlkaneImageBytes.Length);
		$AlkaneMs.Seek(0, [System.IO.SeekOrigin]::Begin)

		#Configure attachment				
		$att1 = new-object Net.Mail.Attachment($AlkaneMs, $ct);
		$att1.ContentType.MediaType = “image/png”;
		$att1.ContentId = "Attachment";
		$att1.ContentDisposition.Inline = $True;
		$att1.ContentDisposition.DispositionType = "Inline";

		#Add attachment to the mail
		$msg.Attachments.Add($att1);

		#get current users email address and display name, based on the USERNAME environment variable		
		$searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
		$currentEmail = $searcher.FindOne().Properties.mail.Item(0).ToString();
		$currentName = $searcher.FindOne().Properties.displayname.Item(0).ToString();

		#Configure email sender, recipients and message
		$MessFrom = New-Object System.Net.Mail.MailAddress $currentEmail, $currentName
		$msg.From = $MessFrom;									
		#$msg.ReplyTo = $currentEmail;
		$msg.To.Add("xxx@yyy.zzz");
		$msg.subject = "Example Email";	
		$msg.IsBodyHTML = $true
		$msg.body = "<div style='font-family:calibri;font-size:10pt; font-weight:bold;'>Alkane Support Team</div>
		<div style='font-family:calibri;font-size:10pt;'>Web: www.alkanesolutions.co.uk</div>
		<div style='font-family:calibri;font-size:10pt;'>Mail: <a href='support@alkanesolutions.co.uk'>support@alkanesolutions.co.uk</a></div><br />
		<img src=cid:Attachment />"				

		#Sending email 
		$smtp.Send($msg)

		[System.Windows.Forms.MessageBox]::Show("Email Sent!" , "Status") 

		#Dispose objects
		$att1.dispose()
		$AlkaneMs.dispose()

	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show($_ , "Status") 				
	}

})
$form.Controls.Add($btn)

#show the form
$drc = $form.ShowDialog()