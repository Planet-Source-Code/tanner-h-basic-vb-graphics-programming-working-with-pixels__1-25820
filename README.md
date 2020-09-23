<div align="center">

## Basic VB Graphics Programming \- working with pixels


</div>

### Description

Many VB programmers don't understand the graphics capabilities of Visual Basic when you use basic API calls. This (and a series of forthcoming tutorials) will explain the basics of getting and setting pixels using both VB and the Windows API (via SetPixel, SetPixelV, and GetPixel), as well as using the API to set pixels on objects other than picture boxes (such as command buttons, frames, etc.).
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tanner H](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tanner-h.md)
**Level**          |Beginner
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tanner-h-basic-vb-graphics-programming-working-with-pixels__1-25820/archive/master.zip)





### Source Code

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<META NAME="Generator" CONTENT="Microsoft Word 97">
<TITLE>Tanner's VB World - Graphics Programming in Visual Basic Tutorial: Setting and Getting Pixels</TITLE>
<META NAME="keywords" CONTENT="Visual Basic, Graphic Programming, Graphics Programming, SetPixel, SetPixelV, GetPixel, API Graphics calls, Tanner, Helland, PSet, Point, Extract RGB, Red, Green, Blue, Tutorial, Information, GetDC, Tanner's VB World">
<META NAME="Version" CONTENT="8.0.4308">
<META NAME="Date" CONTENT="8/15/00">
<META NAME="Template" CONTENT="C:\Program Files\Microsoft Office\Office\Html.dot">
</HEAD>
<BODY TEXT="#000000" LINK="#0000ff" VLINK="#800080" BACKGROUND="tannerhelland.50megs.com/backgrounds/stone.gif">
<B><FONT FACE="Arial" SIZE=4 COLOR="#0000ff"><P>Graphics Programming in Visual Basic - Setting and Getting Pixels </P>
</FONT><FONT FACE="Arial" COLOR="#0000ff"><P>By: Tanner Helland</P><DIR>
<DIR>
</B></FONT><FONT FACE="Arial" SIZE=2 COLOR="#0000ff"><P>Despite what many programmers will tell you, Visual Basic is an excellent programming language for high-end graphic applications. Its easy-to-use interface and programming language allows you to quickly and accurately create all sorts of neat programs without having to worry about the mess of C++ syntax. Also, you can use a number of easy API calls to speed up your interface to professional speed. So, here's part of how to become a professional graphics programmer using only VB.</P>
</FONT><B><FONT FACE="Arial" COLOR="#0000ff"><P>-THE EASY WAY TO DO PIXEL STUFF-</P>
</B></FONT><FONT FACE="Arial" SIZE=2 COLOR="#0000ff"><P>This tutorial will go through the basic way to get and set pixels in Visual Basic. You will use both VB and the Windows API and see the differences between both methods. While this way of getting and setting pixels is slower then the forthcoming part 2 of this tutorial (using GetBitmapBits) it is significantly easier for a beginner, and will still offer impressive results.</P>
<B><P>PART I - GETTING COLORS</P>
</B><P>Before you can do anything to a picture, you have to first get the color of each pixel. There are two intelligent ways to do this, and both are extremely easy.</P>
<P>Way 1 - Using VB</P>
<P>You can use the Point event in VB to get the color of a specified pixel. The format is as follows:</P>
<B><P>Color = PictureBox.Point(x,y)</B> | where PictureBox is the name of the picture box or form you want to retrieve the pixel from, and (x,y) are the pixels coordinates. However, this method is quite slow, and for large pictures it will really start to rack up the time. So basically, don't use it. The best way to get pixels is to use the GetPixel API call:</P>
<P>Way 2 - Using the Windows API</P>
<B><P>Private Declare Function GetPixel lib "gdi32" (ByVal hDC as Long, ByVal x as Long, ByVal y as Long) as Long</P>
<P>Color = GetPixel(PictureBox.hDC, x, y)</B> | where PictureBox is the name of the picture box or form you want to retrieve the pixel from, and (x,y) are the pixels coordinates. This method is many times faster then using VB, and it is basically the same call, except for the API declaration. I will write more on the API call structure in a future tutorial, but for now just trust me. </FONT><FONT FACE="Wingdings" SIZE=2 COLOR="#0000ff">J</FONT><FONT FACE="Arial" SIZE=2 COLOR="#0000ff"> </P>
<B><P>PART II - DRAWING COLORS</P>
</B><P>Just as with getting pixels from a picture box or form, there are several ways to set pixels onto an object as well. Again, the internal VB method is very slow compared to the 2 API calls you can use. For you die-hard VB users, the very slow PSet command is the way to go:</P>
<B><P>PictureBox.PSet (x,y), Color</P>
</B><P>Whereas the API Calls are as follows:</P>
<B><P>Private Declare Function SetPixel lib "gdi32" (ByVal hDC as Long, ByVal x as Long, ByVal y as Long, ByVal Color as Long) as Long</P>
<P>SetPixel PictureBox.hDC, x, y, Color</P>
</B><P>Or:</P>
<B><P>Private Declare Function SetPixelV lib "gdi32" (ByVal hDC as Long, ByVal x as Long, ByVal y as Long, ByVal Color as Long) as Byte</P>
<P>SetPixelV PictureBox.hDC, x, y, Color</P>
</B><P>The only difference between the two functions, if you notice, is that SetPixel returns a Long (the color that the function was able to set) while SetPixelV returns a byte (whether or not the pixel was set). I would always recommend using SetPixelV, simply because it is slightly faster then SetPixel, but the difference is not very noticeable. So, you should now be able to quickly get and set pixels from any picture box or form, right? But, as always, I have some fun things you can add to the useless programming knowledge section of your brain (heh heh).</P>
<B><P>PART III - DRAWING AND GETTING COLORS FROM "SPECIAL" THINGS</P>
</B><P>Up until this point we've been relegated to using only picture boxes and forms because they're the only things that have an accessible hDC property, right? Well, there are certain ways to get around that so that we can set pixels on, say, a command button or a check box. To do this, we use the magical 'GetDC' API call:</P>
<B><P>Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long</P>
<P>Dim TemporaryHandle as Long</P>
<P>TemporaryHandle = GetDC(CommandButton.hWnd) <U>OR</U> GetDC(CheckBox.hWnd) <U>OR</U> GetDC(TextBox.hWnd) etc., etc...</P>
</B><P>Now you can have all sorts of fun! Say, for some odd reason, that you want to set pixels on a command button. After using the GetDC call to assign a handle to the command button, you can do the SetPixel or SetPixelV call using the variable that contains the newly created hDC and - presto - you can draw on almost anything! Play with that API call for kicks if you ever get bored - it's kind fun...</P>
</FONT><P> </P></DIR>
</DIR>
<P><A HREF="http://tannerhelland.50megs.com/VBStuff.htm"><FONT FACE="Arial">Back to Tanner's VB World Home</FONT></A></P>
<P><A HREF="http://tannerhelland.50megs.com"><B><FONT FACE="Arial">Visit the homepage of Tanner Helland</B></FONT></A></P></BODY>
</HTML>

