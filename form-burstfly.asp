<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Function RemoveHTML( strText )
    Dim TAGLIST
    TAGLIST = ";!--;!DOCTYPE;A;ACRONYM;ADDRESS;APPLET;AREA;B;BASE;BASEFONT;" &_
              "BGSOUND;BIG;BLOCKQUOTE;BODY;BR;BUTTON;CAPTION;CENTER;CITE;CODE;" &_
              "COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EM;EMBED;FIELDSET;" &_
              "FONT;FORM;FRAME;FRAMESET;HEAD;H1;H2;H3;H4;H5;H6;HR;HTML;I;IFRAME;IMG;" &_
              "INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LI;LINK;LISTING;MAP;MARQUEE;" &_
              "MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OL;OPTION;PARAM;PLAINTEXT;" &_
              "PRE;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STRONG;STYLE;SUB;SUP;" &_
              "TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;U;UL;VAR;WBR;XMP;"

    Const BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"
    
    Dim nPos1
    Dim nPos2
    Dim nPos3
    Dim strResult
    Dim strTagName
    Dim bRemove
    Dim bSearchForBlock
    
    nPos1 = InStr(strText, "<")
    Do While nPos1 > 0
        nPos2 = InStr(nPos1 + 1, strText, ">")
        If nPos2 > 0 Then
            strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1)
	    strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ")

            nPos3 = InStr(strTagName, " ")
            If nPos3 > 0 Then
                strTagName = Left(strTagName, nPos3 - 1)
            End If
            
            If Left(strTagName, 1) = "/" Then
                strTagName = Mid(strTagName, 2)
                bSearchForBlock = False
            Else
                bSearchForBlock = True
            End If
            
            If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                bRemove = True
                If bSearchForBlock Then
                    If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                        nPos2 = Len(strText)
                        nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare)
                        If nPos3 > 0 Then
                            nPos3 = InStr(nPos3 + 1, strText, ">")
                        End If
                        
                        If nPos3 > 0 Then
                            nPos2 = nPos3
                        End If
                    End If
                End If
            Else
                bRemove = False
            End If
            
            If bRemove Then
                strResult = strResult & Left(strText, nPos1 - 1)
                strText = Mid(strText, nPos2 + 1)
            Else
                strResult = strResult & Left(strText, nPos1)
                strText = Mid(strText, nPos1 + 1)
            End If
        Else
            strResult = strResult & strText
            strText = ""
        End If
        
        nPos1 = InStr(strText, "<")
    Loop
    strResult = strResult & strText
    
    RemoveHTML = strResult
End Function


Function sendmail(msg,subj,rec,bcc)
Dim iMsg 
Dim strHTML
Const CdoReferenceTypeName = 1
Dim objBP
set iMsg = CreateObject("CDO.Message")
Set objBP = iMsg.AddRelatedBodyPart(Server.MapPath("#"), "Onclickprod.com", CdoReferenceTypeName)
	objBP.Fields.Item("urn:schemas:mailheader:Content-ID") = "<Onclickprod.com>"
	objBP.Fields.Update

   strHTML = "<!doctype html>"
    strHTML = strHTML & "<html>"
    strHTML = strHTML & "<head>"
    strHTML = strHTML & "<meta charset=""utf-8"">"
    strHTML = strHTML & "<title>onclickprod | webdesign, branding, user interface </title>"
    strHTML = strHTML & "</head>"
    strHTML = strHTML & "<body style=""font-family:verdana,arial; font-size:12px;background-color:white;"">"
	strHTML = strHTML & "<a href=""http://onclickprod.com"">"
    strHTML = strHTML & "<img src=""cid:CE.com"" border=0 "
    strHTML = strHTML & "ALT=""onclickprod.com"">"
	strHTML = strHTML & "</a><br><br>"
	strHTML = strHTML & msg
	strHTML = strHTML & "<p>Regards<br>"
	strHTML = strHTML & "Onclickprod.com</p>"
    strHTML = strHTML & "</body>"
    strHTML = strHTML & "</html>"
With iMsg
	.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
	.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="mail.yourserver.com"
	.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 
	.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1 
	.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")="" 
	.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")="" 
	.Configuration.Fields.Update
    .To = rec 
    .From = "hello@yourserver.com"
	.Bcc = bcc
	.ReplyTo = "hello@yourserver.com"
    .Subject = subj 
    .HTMLBody = strHTML
    .Send
End With
Set iMsg = Nothing
Set objBP = Nothing
sendmail=true
end function

Function ValidEmail(ByVal emailAddress) 
Dim objRegEx, retVal 
Set objRegEx = CreateObject("VBScript.RegExp") 
With objRegEx 
      .Pattern = "^\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b$" 
      .IgnoreCase = True 
End With 
retVal = objRegEx.Test(emailAddress) 
Set objRegEx = Nothing 
ValidEmail = retVal 
End Function


dim nom
nom=""
if request.form("nom")<>"" then nom=RemoveHTML(request.form("nom")) else er=er&"_n"

dim email
email=""
if ValidEmail(request.form("email")) then email=request.form("email") else er=er&"_e"

dim telephone
telephone=""
if request.form("telephone")<>"" then telephone=RemoveHTML(request.form("telephone")) else er=er&"_t"

dim message
contactComment=""
if request.form("message")<>"" then message=RemoveHTML(request.form("message"))

dim objet
objet=""
if request.form("objet")<>"" then objet=RemoveHTML(request.form("objet"))


dim msg
msg="Nouveau contact - immobilier !!<br><br>"
msg=msg&"Nom : "&nom&"<br>"
msg=msg&"Email : "&email&"<br>"
msg=msg&"Telephone : "&telephone&"<br>"
msg=msg&"Objet : "&objet&"<br>"
msg=msg&"Message : "&message
subj="Contact Edimm  - Immobilier"
rec="lacostexvr@gmail.com"
bcc="lacostexvr@gmail.com"
if er="" then 
a=sendmail(msg,subj,rec,bcc)
response.write "<span style=""color:#43ADE3"">Thank You! M."&nom&", your message has been sent !</span>" 
else 
if instr(er,"_n")>0 then mssage="nom"
if instr(er,"_e")>0 then mssage="email"
if instr(er,"_t")>0 then mssage="telephone"
if instr(er,"_n_e_t")>0 then mssage="nom, email et telephone"
response.write "Erreur ! Merci de remplir l'onglet avec votre "&mssage
end if
%>