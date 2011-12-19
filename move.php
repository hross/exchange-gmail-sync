<?php
date_default_timezone_set('America/New_York');
include("ExchangeClient/Exchangeclient.php");
?>

<?php
//connect to imap (gmail account details here)
$mbox = "{imap.gmail.com:993/imap/ssl}";
$conn = imap_open($mbox, "myemail@gmail.com", "password") or die("can't connect: " . imap_last_error());

try {
	// connect to exchange
	$exchangeclient = new Exchangeclient();
	
	// enter exchange user name and password here
	$exchangeclient->init("My Exchange Account", "password");


	// get message from exchange
	$messages = $exchangeclient->get_messages();

	for($i = sizeof($messages)-1; $i >= 0; $i--) {
		$message = $messages[$i];
		echo("Importing message with subject: $message->subject");

		$imsg = exchange_to_imap($message);

		// create message on imap server
		if (imap_append($conn, $mbox, $imsg) === false) {
			imap_append($conn, $mbox, error_email());
			die( "could not append message: " . imap_last_error());
		}
		
		var_dump($message);
		// if we made it, remove the message
		$exchangeclient->delete_message($message->ItemId);

		if (12 == $i) break; // testing, don't loop through everything
	}
} catch (Exception $e) {
	imap_append($conn, $mbox, error_email());
}

imap_close($conn);

function error_email() {
	$envelope["from"]= "error@error.com";
	$envelope["to"]  = "error@error.com";
	$envelope["subject"] = "Problem transferring email";
	
	$part["type"] = TYPETEXT;
	$part["subtype"] = "plain";
	$part["contents.data"] = "Could not transfer data!";
	
	$body[1] = $part;
	return imap_mail_compose($envelope, $body);
}

function exchange_to_imap($message) {
	$to = array();
	$toaddress = "";
	if(isset($message->to_recipients)) {
		for($i = 0; $i < sizeof($message->to_recipients); $i++) {
			$recipient = $message->to_recipients[$i];
			$recip = new stdClass();
			$eparts = explode("@", $recipient->EmailAddress);
			$recip->personal = $recipient->Name;
			$recip->mailbox = $eparts[0];
			$recip->host = $eparts[1];
			$toaddress .= "\"" . $recipient->Name . "\" <" . $recipient->EmailAddress . ">";
			if($i < (sizeof($message->to_recipients)-1)) $toaddress .= ", ";
			$to[] = $recip;
		}
	}

	$cc = array();
	$ccaddress = "";
	if(isset($message->cc_recipients)) {	
		for($i = 0; $i < sizeof($message->cc_recipients); $i++) {
			$recipient = $message->cc_recipients[$i];
			$recip = new stdClass();
			$eparts = explode("@", $recipient->EmailAddress);
			$recip->personal = $recipient->Name;
			$recip->mailbox = $eparts[0];
			$recip->host = $eparts[1];
			$ccaddress .= "\"" . $recipient->Name . "\" <" . $recipient->EmailAddress . ">";
			if($i < (sizeof($message->cc_recipients)-1)) $ccaddress += ", ";
			$cc[] = $recip;
		}
	}
	
	$bcc = array();
	$bccaddress = "";
	if(isset($message->bcc_recipients)) {
		for($i = 0; $i < sizeof($message->bcc_recipients); $i++) {
			$recipient = $message->bcc_recipients[$i];
			$recip = new stdClass();
			$eparts = explode("@", $recipient->EmailAddress);
			$recip->personal = $recipient->Name;
			$recip->mailbox = $eparts[0];
			$recip->host = $eparts[1];
			$bccaddress .= "\"" . $recipient->Name . "\" <" . $recipient->EmailAddress . ">";
			if($i < (sizeof($message->bcc_recipients)-1)) $bccaddress += ", ";
			$bcc[] = $recip;
		}
	}

	$envelope["subject"] = $message->subject;
	$envelope["date"] = date('r', strtotime($message->time_sent));
	$envelope["from"]= "\"" . $message->from_name . "\" <" . $message->from . ">";
	
	if($to) { $envelope["to"] = $toaddress; }
	if($cc) { $envelope["cc"] = $ccaddress; }
	if($bcc) { $envelope["bcc"] = $bccaddress; }

	if ($message->bodytype == 'Text') {
		$part["type"] = TYPETEXT;
		$part["subtype"] = "plain";
		//$part["description"] = "part description";
		$part["contents.data"] = $message->bodytext;
	} else if ($message->bodytype == 'HTML') {
		//html
		$part["type"] = TYPETEXT;
		$part["subtype"] = "html";
		//$part["description"] = "part description";
		$part["contents.data"] = $message->bodytext;
	} else {
		//unknown!
		return NULL;
	}
	
	// append attachments with message->attachments
	if($message->attachments) {
		$part1["type"] = TYPEMULTIPART;
		$part1["subtype"] = "mixed";
		$body[1] = $part1;
		
		// go through attachments and add as parts
		$i = 2;
		foreach($message->attachments as $attachment) {
			$apart["type"] = TYPEAPPLICATION;
			$apart["encoding"] = ENCBINARY;
			$apart['disposition.type'] = 'attachment';
            $apart['disposition'] = array ('filename' => $attachment->Name);
            $apart['dparameters.filename'] = $attachment->Name;
            $apart['parameters.name'] = $attachment->Name;
			$apart["subtype"] = "octet-stream";
			$apart["description"] = $attachment->Name;
			$apart["contents.data"] = $attachment->Content;
			$body[$i] = $apart;
			$i++;
		}
		
		$body[$i] = $part; // add message 
	} else {
		// no attachments
		$body[1] = $part;
	}
	
	// give us back the mail message
	return imap_mail_compose($envelope, $body);
}
?>
