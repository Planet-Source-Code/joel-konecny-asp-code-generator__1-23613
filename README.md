<div align="center">

## ASP Code Generator


</div>

### Description

the engine i am providing is an asp code generator that i am currently using on my website (www.intratelligent.com) to allow asp developers to upload access databases online and build simple asp interfaces to interact with their data. (insert / update / delete and reporting functionality).
 
### More Info
 
the only input that the code requires is an access database and the base path for the project... you will notice that all parameters are left variant because this was originally built as a dll for calls from asp.

there are a few dependencies that my code relies on... dao 3.6 and microsoft scripting runtime (filesystemobject)... the reason i am using dao is for database information (adox in next release). the vb version also uses the common dialog component (standard with vb).

the main function will return a boolean value letting you know whethere the function succeeds or fails... the vb version let's you know with a msgbox.

the code is completely independant of anything on your system... once the code is generated it can be transfered to any iis server and run as long as the server supports asp and ado connectivity.


<span>             |<span>
---                |---
**Submitted On**   |2001-05-31 00:57:24
**By**             |[Joel Konecny](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joel-konecny.md)
**Level**          |Advanced
**User Rating**    |4.9 (64 globes from 13 users)
**Compatibility**  |VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[ASP Code G203145312001\.zip](https://github.com/Planet-Source-Code/joel-konecny-asp-code-generator__1-23613/archive/master.zip)








