Title: Client/Server Authentication Example (winsock)
Description: This example provides a good way of how to securly authenticate your client programs when connecting to your server. This prevents the use of 3rd party clients and unauthorized connections (such as telnet connections and 'reverse engeneering').
How it works: Once a client connects, the server transmits an encrypted string of 100 random characters (You may change this) and the client must respond with the decrypted version. The advantage here is, the server will send a DIFFERENT encryption string for each new connection! So if a 'hacker' finds out what code the client sent back, it will be of no use to him as when a new connection is made, a different 'key' is required to be sent back!
This program also supports the authentication of more than one connection at a time.
I would appreciate votes and comments, and would love to know if you use this method in any program you make.
Thanks! :D
This file came from Planet-Source-Code.com...the home millions of lines of source code
You can view comments on this code/and or vote on it at: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=36626&lngWId=1

The author may have retained certain copyrights to this code...please observe their request and the law by reviewing all copyright conditions at the above URL.
