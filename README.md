Ticket System

This is a simple ticket system where you can submit and read submitted tickets. It uses a share (for example \\192.168.0.1\tickets) to be able to distribute the tickets with other computers. Therefore, it is important that the system has access to this distribution and the users who submit.

The ticket system consists of two different applications; ReadTickets and Maketicket. Just as the names suggest, MakeTicket is used to submit tickets and ReadTickets is used to read and manage them.

![image](https://github.com/user-attachments/assets/c4ff919a-456d-4ada-b60e-a07f5b91f743)

In settings you can set where it should retrieve its tickets (of course in the same place they are sent from MakeTicket) but you can also choose users. Of course you have to choose where to retrieve your tickets for it to work correctly. In some cases you need to restart it.

![image](https://github.com/user-attachments/assets/e4f0a409-653c-455f-9231-a7772528bb49)


The path.txt file is for MakeTicket and must be placed in the same folder. This is how you tell the application where to place the new tickets.



More admin users

To create users for the Ticket System, you must manually (for now) edit a json file (for example \\192.168.0.1\tickets\owners.json), with the format:

{

"ticketOwners": "User1, User2,..."

}

and save it. There will probably be an update in the future.

![image](https://github.com/user-attachments/assets/54743df9-6d17-4a68-ad8a-10e82265ed5e)



Compiled version

I also provide a compiled version in exe format for those who are interested. It is created using the ps2exe tool by Markus Scholtes (https://github.com/MScholtes/PS2EXE). 

Since I don't use any certificates when building the applications, you will probably get a message that it is not secure. You run it at your own risk, of course, and if you are unsure and don't want to take the risk, you can build it yourself using the PowerShell files in this repo (ReadTickets.ps1 and MakeTicket.ps1).
