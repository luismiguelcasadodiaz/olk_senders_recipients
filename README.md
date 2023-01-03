# olk_senders_recipients
 VBA Macro navigates folders extracting from each message the sender and recipients to excel
 
 You need to copy paste this script into the visual basic IDE you get from OUtlook developer menu.
 
 The script walks folders hierarchy and nurtures two excel sheets
 
 Sheet Messages
 

msgCode	msgFolder	                            msgSender	            msgReceived	        msgSent To	                                                               msgCC	msgBCC	msgSubject	                    msgRecipient
0000000	root\luis@fruitspec.com\Deleted Items	gal.peer@fruitspec.com	03/02/2022 12:45	Luis Miguel Casado Diaz; Nadav Leshem			                                            Re: Spain Demo 	                luis@fruitspec.com
0000000	root\luis@fruitspec.com\Deleted Items	gal.peer@fruitspec.com	03/02/2022 12:45	Luis Miguel Casado Diaz; Nadav Leshem			                                            Re: Spain Demo 	                Nadav@fruitspec.com
0000001	root\luis@fruitspec.com\Deleted Items	Raviv@fruitspec.com	    02/02/2022 16:03	Nadav Leshem; Victor Muñoz; Deon Pelser; Luis Miguel Casado Diaz; Gal Peer			        Re: Sales board 	            Nadav@fruitspec.com
0000001	root\luis@fruitspec.com\Deleted Items	Raviv@fruitspec.com	    02/02/2022 16:03	Nadav Leshem; Victor Muñoz; Deon Pelser; Luis Miguel Casado Diaz; Gal Peer			        Re: Sales board 	            Victor@fruitspec.com
0000001	root\luis@fruitspec.com\Deleted Items	Raviv@fruitspec.com	    02/02/2022 16:03	Nadav Leshem; Victor Muñoz; Deon Pelser; Luis Miguel Casado Diaz; Gal Peer			        Re: Sales board 	            Deon@fruitspec.com
0000001	root\luis@fruitspec.com\Deleted Items	Raviv@fruitspec.com	    02/02/2022 16:03	Nadav Leshem; Victor Muñoz; Deon Pelser; Luis Miguel Casado Diaz; Gal Peer			        Re: Sales board 	            luis@fruitspec.com
0000001	root\luis@fruitspec.com\Deleted Items	Raviv@fruitspec.com	    02/02/2022 16:03	Nadav Leshem; Victor Muñoz; Deon Pelser; Luis Miguel Casado Diaz; Gal Peer			        Re: Sales board 	            gal.peer@fruitspec.com
0000002	root\luis@fruitspec.com\Deleted Items	anamtrifu@gmail.com	    27/01/2022 18:45	Luis Miguel Casado Diaz			                                                            Fwd: Project definition est2XXX	luis@fruitspec.com
0000003	root\luis@fruitspec.com\Deleted Items	book3@escodealmaker.com	20/01/2022 00:49	Luis Miguel Casado Diaz			                                                            Would you mind  participate 0	luis@fruitspec.com
0000004	root\luis@fruitspec.com\Deleted Items	sadnot@satsintesis.com	21/12/2021 10:44	Luis Miguel Casado Diaz			                                                            RE: Felicitación de Navidad	    luis@fruitspec.com


 
 Sheel SMPT
 
It is a list of unique smtp Address. It is the seed table of a further smtp clasification 
[Internal/external]  
Internal ==> [comercial/finances/Operations/R&D/IT]
External ==> [Customer, Lead, Supplier]


smtpAddress						smtpDomain
luis.casado@iese.net			iese.net
luismiguelcasadodiaz@gmail.com	gmail.com


Before executing the Script be sure outlook setting Mail to keep offline (locally) ALL messages, not only last year messages (default settings)