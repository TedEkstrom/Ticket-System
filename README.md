Ticket System

This is a simple ticket system where you can submit and read submitted tickets. It uses a distribution to be able to distribute the tickets with other computers. Therefore, it is important that the system has access to this distribution and the users who submit.

The ticket system consists of two different applications; ReadTickets and Maketicket. Just as the names suggest, MakeTicket is used to submit tickets and ReadTickets is used to read and manage them.

![image](https://github.com/user-attachments/assets/c4ff919a-456d-4ada-b60e-a07f5b91f743)

![image](https://github.com/user-attachments/assets/d27cda0b-5f3c-40e1-8f44-031bbee6bc2c)



More admin users

To create users for the Ticket System, you must manually (for now) edit a json file with the format:

{

"ticketOwners": "User1, User2,..."

}

and save it with the file name owners.json and which is placed in the root of your tickets path, for example \\192.168.0.1\tickets\owners.json. There will probably be an update in the future.

![image](https://github.com/user-attachments/assets/54743df9-6d17-4a68-ad8a-10e82265ed5e)



Compiled version

I also provide a compiled version in exe format for those who are interested. It is created using the ps2exe tool by Markus Scholtes (https://github.com/MScholtes/PS2EXE).
