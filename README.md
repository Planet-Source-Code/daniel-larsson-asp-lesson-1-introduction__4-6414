<div align="center">

## ASP Lesson \#1 \- Introduction


</div>

### Description

This is the first lesson of many upcoming.<p>I will learn you what ASP is, and what to use it for in this lesson. You will not be tought any ASP Code the first lesson. This is just what it says it is.....an Introduction.<p>If you're an experienced developer, you won't get much of reading this lesson.<p>I will also tell you where to get a good Reference, and a few things about ASP Developement languages.<p>Please rate it, even if you don't like it, and tell me what can be better and what is bad.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Larsson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-larsson.md)
**Level**          |Beginner
**User Rating**    |4.3 (94 globes from 22 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Documents/ Frames](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/documents-frames__4-27.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-larsson-asp-lesson-1-introduction__4-6414/archive/master.zip)





### Source Code

```
<hr><b>In this lesson and many upcomming lessons you will eventually learn everything about ASP. In this lesson we will not go any further than try to explain how I plan to write this guide, what ASP is, where you find References, and of course the most basics of ASP Coding.</b>
<hr>
<b>What is ASP??</b>
<p> ASP stands for Active Server Pages. This means that all code is generated on the server. In other words, it doesn't matter which browser you have, it doesn't matter which version you have or manufacturer. Everything is generated in the serve's software. When it's going to be showned, it prints it out in simple HTML-Code which will be seen if the user access "View Source".
<p>
If the user can't see the code, where do you then put the code? In a file! When the visitor wants to visit a page that contains ASP-Code the following procedure is made:<p><table border="0">
<tr><td>1.</td><td>The visitor demands the page.</td>
<tr><td>2. </td><td>The file is generated through the server's software, which reads the ASP-Code and execute what it asks for.</td>
<tr><td>3. </td><td>The code which now, hopefully generated something, deletes itself from the file and everything that is showned in the visitor's Webbrowser is a simple HTML-Document.</td></table>
<p>
The code, or script, which is more correct usually demands special softwares installed on the server in order for things to work. The most common is that the server is running Microsoft Internet Information Server as Software ( IIS ). This is microsoft's easy distribution of this software. Microsoft Personal Webserver has limited functions compared to IIS, But it's enough to develope ASP for a common PC-Computer.<p>
ASP executes, like i said earlier, at the server page, but there is no Code which is called ASP. All code that is written is called VBScript(Visual Basic Script) or Jscript(Java Script). You can use any of these codelanguages to produce an ASP Page. But in this guide we will write about VBScript only. VBScript is a shorter version of Microsoft's real Visual Basic Language.<br><hr><br><b>What can I use ASP for?</b><p>This is a question that has no real answer, I would say it's only your imagination that is your enemy ;). An answer like that is hard to understand for someone that hardly knows what ASP is. I'm therfor trying to explain it a little more. ASP can be used for this:<p><table border="0"><tr><td>*</td><td>Store and get information from a database from the Internet.</td><tr><td>*</td><td>Make your webpages more interactive for your visitors. A golden example of this is for example, show the current date and time at your page when anyone visits it.</td><tr><td>*</td><td>Make your pages more personal, by letting the visitor choose the Fonts and size he wants on the page.</td><tr><td>*</td><td>Store information about a visitor which can be used everytime that visitor comes back. And because of that remind him about something or show him his personal settings for your page.</td><tr><td>*</td><td>Make games, Guestbooks, Advertisements, Create and read cookie information etc.</td></table><p>
Think of applications like Guestbooks, Advertisement pages, Alertpages, News pages and chatpages are all based on the same code. Only a few cosmic changes was made. A guestbook often works in a way like this: A user enters an information in a form. Than that information is stored in a long file together with all the other users information.<p>
The advertisement page works in the same way. The users enters a text and than his text is showned together with all other texts that have been stored in a file or a database. Both the code and the princip is the same, only a few modification was made.<br><hr><br><b>What will you learn?</b><p>Now we are coming to the most important of all. What will I learn?? I will teach you how to produce ASP pages in the VBScript Language. VBScript is a language exactly like many others. This means there are many different categories where you can place all the different commands which you can use.<p>Mainly it's these categories:<p><table border="0"><tr><td>*</td><td>Functions</td><tr><td>*</td><td>Objects</td><tr><td>*</td><td>Methods</td><tr><td>*</td><td>Operators</td><tr><td>*</td><td>Properties</td><tr><td>*</td><td>Statements</td></table><p>
These categories contains many different commands which works and are used in different ways. Alls the commands are most likly used together with other commands from different categories. Therefor we will jump a little between them just to teach you the basics and the most important things. But of course, Later we will go deeper into the commands, methods and objects. In the beginning I will teach you the differents between the categories.<br><hr><p><b>For those who wants to try self!</b><p>If you believe that all the commands that are used in VBScript is stored in the programmer's brain.....your WRONG! There are very many, if not endless, different commands and categories. It's about impossible for anyone to keep all this in his head. Therefor there are something called References to help you.<p>The difference between Guides, like this, and references is that you dont tell anything in the right way. They only show you how to use it, no information of problems, where to use it and such things. There are only short examples on how to use them. For those that have never ever before coded before, this is like Chineesee!!! IF you need a reference theres a good one at Microfofts homepage at: <a href="http://www.microsoft.com/scripting/vbscript/">http://www.microsoft.com/scripting/vbscript/</a>.<br><hr><br><b>Databases, Databases, when will we ever use databases??</b><p>At a later stage we will learn how to work with a database, but think about that a database have many different ways of beeing used. This means that there are certain things that you need to know about VBScript to have any use of your database. But don't get scared, VBScript is very Easy!!!<br><hr><br><b>Some basics about ASP in the first lesson</b><p>Let's go through some very basic things about ASP.<p>&lt;html&gt;<p>&lt;% ASP Code %&gt;<p>&lt;/html&gt;<p> These lines are with a % in front or after itself are called "Blockers". This is where the code starts and stops. All ASP code must be written within these tags ( <% and %> ). If they don't stand within these tags, the code will not be executed and therfore not showned.<p>'Here is a comment<p>As soon as you would like to make a little comment to your code you make a '. Everything write to the right of the ' will not be executed in the code. You dont have to make anything to make the comment tag dissapear. If you want a comment over two rows do like this:<p>'comment 1<br>'comment 2<p>-<p>This is an underscore, this is used whenever you want your code to continue on the next row. This is usefull when you want to break a long row of code into two lines. Just use a Underscore (_) and continue on the next line. Example:<p>Dim Sak,_<br>Sak2<p>This is only made to make the code look easiser and much simplier to understand for the programmer. At least that's the mainreason to use an underscore.<p>Hope that you continue reading my ASP Guides, since in the next lesson I help you put up a Webserver for your ASP Files.<p><b>End of lesson #1!</b><hr>
```

