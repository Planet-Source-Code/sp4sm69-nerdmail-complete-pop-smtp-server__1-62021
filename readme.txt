Welcome to the NerdMail POP/SMTP server! Please read this document from start to finish before beginning, as it is meant as an introductory tutorial.

You have just downloaded the ONLY fully functional Visual Basic/MySQL mail server! It contains a POP server, an SMTP server, and a webmail interface, all contained inside one application. The server runs in the system tray, and the interface is available by double clicking the icon.

The easiest way to use NerdMail is by compiling it before anything else. The first time you run NerdMail, it will run its built-in setup wizard. The setup wizard will ask a few questions to determine how the server should act.

  <>POP Server IP        : A combo box is shown with all available IP interfaces. Select one for the POP server
  <>SMTP Server IP       : A combo box is shown with all available IP interfaces. Select one for the SMTP server
  <>Root Domain          : The domain for which your server is providing the service. Will follow the @ in email addresses

  <>MySQL Server         : The IP address or domain of the MySQL server on which your database is located
  <>MySQL Username       : The username used to connect to the MySQL Server
  <>MySQL Password       : The password used to connect to the MySQL Server
  <>MySQL Database       : The MySQL database used
  <>Database needs setup : Creates the database (if sufficient priveleges), and creates the tables

  <>Setup first user     : The setup wizard will create the first user for you
  <>Username             : The username of the first user (username@domain.tld)
  <>Password             : The password of the first user
  <>Verify Password      : Verify the password entered

Click ok, and the setup wizard will save your settings. To test the server, do the following:

-Compile (if you haven't already) the mail server, and run it.
-Open your webbrowser and go to http://domain:81, substituting your chosen domain into the address.
-If you did not create a used in the setup wizard, click signup and fill out the form.
-Head to your inbox. You should have a welcome email in your inbox.
-Open your preferred email client, setup your account, and check your email
-Send an email from your client to yourself, username@domain, then check your email. It may take several seconds to send.
-Send an email to someone on some other email server, and determine if the mail arrived. Again, it may take several seconds.
-Go to another email service (such as Hotmal or Gmail) and try to send an email to your account. This step will often fail, because ISPs often block incoming connections on port 25. In order for the mail server to work, the computer must be able to accept connections on port 25.
-Send an email to kingofgeeks1@yahoo.com and say "Your mail server is working!"
-Vote for this code at www.pscode.com

And there you have it, a complete POP/SMTP Mail server. If any of the above steps fail, do not hestitate to contact me for help. I will gladly assist in any way that I can.

Please note at that this server comes with no warranty or guarantee. I take no responsibility if your hand gets eaten off by a lawn mower, your cup of McDonald's coffee burning you to death when it spills, crashing of the internet, any unexplainable spontaneous cobustion, mild cases of death, or any other such consequences. Basically, you get you pay for, and this was free.

Last (but most definitely NOT least), I would like to thank Ashley Harris. His mail server proved to be a great help (not to mention the inspiration for this project), and also gave some very good code (which he has given me permission to use). Without his mail server, we would not have this mail server.

King of Geeks
kingofgeeks1@yahoo.com
Y! KingOfGeeks1
AIM: Steve0Bob666

PS) I apologize for lack of commenting in some (or most XD) parts of the code. If you would like explaination for certain parts, contact me and I'll be glad to help.