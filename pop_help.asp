<html>
	<head>
		<title>HELP</title>
		<link rel=stylesheet HREF="includes/all.css">
	</head>
<Body bgcolor="black" link="white" alink="white" vlink="white" text="#ffffff">
<font face="arial" size="2">
<%
select case Request.QueryString("ID")
	case 1
%>
Overview:
<BR />
This is a quick glance at your nations status, covering how many units you have, if your continent it under attack, if you yourself have any incomming units and details on your research and construction teams.
<%
	case 2
%>
Private News:
<BR />
This holds all the current entries about your nation and events that have happened to it.  Everytime something noteworthy happens like your attacked or research is compleated you will get an entry in your private news.  You can delete all the enteries in your private news by using the delete button at the bottom of the page.
<%
	case 3
%>
Continental Status:
<BR />
This tells you about all the incomming and outgoing forces that are affecting your continent.  This is espically useful for orgnising defence.  
<%
	case 4
%>
Messages (inbox):
<BR />
This page lists all the messages you have in your inbox from other players.  All the new messages you have are highlited in green.  The buttons at the bottom of the page allow you to delete your unwanted mail and send new messages.  
<%
	case 5
%>
Messages (read message):
<BR />
This page allows you to read and reply to messages you have recieved.
<%
	case 6
%>
Messages (new and reply):
<BR />
This page allows you to create new messages and type your reply to messages you wish to reply to.  Please remember that being rude/offensive is cause for being deleted.
<%
	case 7
%>
Continental Forum:
<BR />
This page allows you to decuss issues with the players that are in the same continent as you, you can make new topics or reply to existing ones.  
<%
	case 8
%>
Explore:
<BR />
This page allows you to train explorers then use them to discover more land for your nation.  The more land that you explore the higher the cost will be in explorers so exploration can become a very expensive business.
<%
	case 9
%>
Development:
<BR />
This page allows you to use the land you have to build industry on.  Each industrial unit will produce 100 of its resource each tick and you will get a base of 300 of each resource each tick.  Remember your industry can be destored when your attacked.  
<%
	case 10
%>
Research:
<BR />
This page allows you to progress down the research section of the technology tree.  Simply choose which research you would like to start then click the start button at the side.  Remember you can only research one thing at a time and the resources required are deducted immidiatly.  When the research is finishes you will get an entery in your private news to let you know.
<%
	case 11
%>
Construction:
<BR />
This page allows you to process down the construction section of the technology tree.  Simply choose which constrctuon you would like to start then click the stasrt button at the side.  Remember you can only build one thing at a time and the resourced required are dedeucted immidiatly.  When the construction finishes you will get an entry in your private news to let you know.
<%
	case 12
%>
Military Units:
<Br />
This page allows you to order more and track the production of the units you will use to attack and defend with.  The units you have avaible to use is limited by your progress through the technology tree.  Once ordered units can not be canceled.  
<%
	case 13
%>
Defence Network:
<BR />
This page allows you to order the defence units which only defend you while being attacked.  These units are cheeper than the equliviant mobile units but they can not move.  The units you can order is limited by your progess through the technology tree.  Once ordered units can not be canceled. 
<%
	case 14
%>
Military Opperations:
<BR />
This page allows you to control what your militry forces does.  The first section of the page allows you to send out forces to attack and defend other nations.  Each unit has an ETA and the force you sent will all arrive at the same time, taking the highest ETA to arrive at the target country.  Each unit willl try to target each unit to the extent it can be done, for excample, infantry will try and target tanks as they could destroy them but will not target bombers as they are too high to be damanged.  The lower half of the page is used to view and recal the forces that you have traveling to and from other nations.
<%
	case 15
%>
Special Opperations Execetive:
<BR />
This page allows you to perform less government friendly opperations.  Here you can trains your spies for the SEO and your MP's to defend yourself from other countries SEO opperations.  When sending spies out on missions the more spies you send the higher the chance of success, but the more you put at risk of being captued and killed.  If your spies are captures your target will foind out you have been performing SEO opperations on them.  
<%
	case 16
%>
Polotics:
<BR />
This is the page in which you vote for your power base nation for your contient.  The nations are shown in order of the number of votes they have for them.  If you are the power base nation then you can change the continent options from this page.
<%
	case 17
%>
Contient:
<BR />
This page is used to view other continents.  The banner at the top is the continent banner and the nations are displayed below.  This page is mainly used when looking for attack targets.  
<%
	case 18
%>
World:
<BR />
This page allows you to see the top 10 largest and richest nations, these are the top players in the game and probally have the most military power.
<%
	case 19
%>
Settings:
<BR />
This page allows you to change your account settings, if you wish to go into holiday mode (stops your account being deleted for inactivity and you being attacked and can only be used twice) then you can do that here.  If you do not wish to change your password then leave the password fields blank.
<%
end select
%>
<BR />
<BR />
<A HREF="javascript:window.close()">Close Window</A>
</font>
</body>
</html>