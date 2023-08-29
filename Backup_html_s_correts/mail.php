<?php
// Criando a variável que irá representar o corpo do e-mail
$msg = "De: ";
$msg .= $Nome;
$msg .= "\r\n";
$msg .= $assunto ='Contato direto do site www.acessoriosbayer.com.br';
$msg .= "\r\n";
$msg .= "Empresa: ";
$msg .= $Empresa;
$msg .= "\r\n";
$msg .= "Telefone: ";
$msg .= $Telefone;
$msg .= "\r\n";
$msg .= "Email: ";
$msg .= $Email;
$msg .= "\r\n";
$msg .= "Motivo do contato: ";
$msg .= $Motivo;
$msg .= "\r\n";
$msg .= "Mensagem: ";
$msg .= $Mensagem;
$msg .= "\r\n";
//$msg .= "Assunto:";

//$msg .= "Mensagem:";
//$msg .= $mensagem;
//$msg .= "\r\n";
$msg .= "\r\n";
// Criando a variável adicional de headers
$headers = "From: "; //Observe que eu utilizei o header 'From' que é um header padrão.
$headers .= $nome;
$headers .= "<";
$headers .= $email;
$headers .= ">";
// Agora é só utilizar a função mail()
// mail("contato@meusite.com", $assunto, $msg, $headers);
// Na linha de comentário acima eu mostro como ficaria a função.
// O bloco exibe a mensagem de acordo com resultado da utilização
// da função, ou seja, se a mensagem foi enviada ou não
if (@mail("contato@acessoriosbayer.com.br", $assunto, $msg, $headers)) {
?>
<%@LANGUAGE="VBSCRIPT"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0031)http://www.acessoriosbayer.com.br/ -->
<% SESSION.LCID = 1046  %><!--#include file="Connections/conexao.asp" --><%
Dim rs_notici
Dim rs_notici_numRows
Set rs_notici = Server.CreateObject("ADODB.Recordset")
rs_notici.ActiveConnection = MM_conexao_STRING
rs_notici.Source = "SELECT * FROM noticias ORDER BY data DESC"
rs_notici.CursorType = 0
rs_notici.CursorLocation = 2
rs_notici.LockType = 1
rs_notici.Open()
rs_notici_numRows = 0
%><HTML><HEAD><TITLE>BAYER - ACESS&Oacute;RIOS P/ VIDROS TEMPERADOS</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<STYLE type=text/css>
BODY {
	BACKGROUND-IMAGE: url(img/body_bg.jpg); MARGIN: 0px; BACKGROUND-REPEAT: repeat-x; BACKGROUND-COLOR: #FFFFFF}
A:link {
	COLOR: #333333}
A:visited {
	COLOR: #333333}
A:hover {
	COLOR: #FF6600}
A:active {
	COLOR: #333333}
.style29 {
	font-family: ARIAL;
	font-weight: bold;
	color: #333333;
	font-size: 14px;
}
.style40 {color: #000000}
.style43 {
	font-family: ARIAL;
	font-weight: bold;
	color: #FF3300;
}
.style44 {color: #333333}
.style45 {color: #FF6600}
</STYLE>

<SCRIPT type=text/javascript>
_uacct = "UA-371279-1";
urchinTracker();
</SCRIPT>

<META content="MSHTML 6.00.6000.17037" name=GENERATOR></HEAD>
<BODY>
<TABLE cellSpacing=0 cellPadding=0 width=770 align=center border=0>
  <TBODY>
  <TR>
    <TD align="center"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="865" height="251">
      <param name="movie" value="Flash/menu_bayer.swf">
      <param name="quality" value="high">
      <embed src="Flash/menu_bayer.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="865" height="251"></embed>
    </object></TD>
  </TR>
  <TR>
    <TD height=413 align=left valign="top" borderColor=#333333 
    background=img/body_bg_home.jpg 
    bgColor=#ffffff><p>&nbsp;</p>
      <table width="836" height="380" border="0">
        <tr>
          <td width="635" height="376" align="center" valign="top"><table width="576" height="520" border="0">
            <tr>
              <td width="566" height="56" align="left" valign="top"><div align="left"><img src="img/home_produtos/contato_topic.jpg" width="166" height="18"></div></td>
            </tr>
            <tr>
              <td valign="top"><table height="211" width="482" border="0">
                <tbody>
                  <tr>
                    <td valign="top" align="center" bgcolor="#FFFFFF" height="207"><table width="310" height="137" border="0">
                      <tr>
                        <td><DIV align="center" class="style40">
                            <p class="style29">MENSAGEM ENVIADA COM SUCESSO! </p>
                            <p class="style29 style45">AGUARDE CONTATO. </p>
                            <p><span class="style43"><a href="index.html" class="style43">voltar</a></span></p>
                        </DIV></td>
                      </tr>
                    </table></td>
                  </tr>
                </tbody>
              </table>                <p><br>
              </p>                </td>
            </tr>
            <tr>
              <td height="21" valign="top">&nbsp;</td>
            </tr>
            
            
          </table></td>
          <td width="191" align="center" valign="top"><p>&nbsp;</p>
            <p>&nbsp;</p>
            <table width="200" border="0">
            <tr>
              <td width="12" height="43"><img src="img/home_produtos/set.jpg" width="9" height="12"></td>
              <td width="163" align="left"><div align="left"><a href="kitsfechados.html" class="style29 style44">KITS FECHADOS </a></div></td>
            </tr>
            <tr>
              <td height="40"><img src="img/home_produtos/set.jpg" width="9" height="12"></td>
              <td align="left" class="style29"><div align="left"><a href="acess_avuls.html">ACESSÓRIOS AVULSOS</a></div></td>
            </tr>
            <tr>
              <td height="39"><img src="img/home_produtos/set.jpg" width="9" height="12"></td>
              <td align="left"><div align="left"><a href="roldanas_carrinhos.html" class="style29">ROLDANAS E CARRINHOS</a></div></td>
            </tr>
            <tr>
              <td height="35"><img src="img/home_produtos/set.jpg" width="9" height="12"></td>
              <td align="left"><div align="left"><a href="roldanas_carrinhos.html" class="style29">PUXADORES</a></div></td>
            </tr>
            <tr>
              <td height="37"><img src="img/home_produtos/set.jpg" width="9" height="12"></td>
              <td align="left"><div align="left"><a href="rold_sacada.html" class="style29">KIT ROLD. P/ SACADA</a></div></td>
            </tr>
            <tr>
              <td height="34"><img src="img/home_produtos/set.jpg" width="9" height="12"></td>
              <td align="left"><div align="left"><a href="roldana_pia_vitrine.html" class="style29">ROLD. P/ PIA E VITRINE </a></div></td>
            </tr>
            <tr>
              <td height="21">&nbsp;</td>
              <td align="left">&nbsp;</td>
            </tr>
            <tr>
              <td height="21">&nbsp;</td>
              <td align="left">&nbsp;</td>
            </tr>
          </table></td>
        </tr>
      </table>
      </TD>
  </TR>
  <TR>
    <TD align=center valign="top"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="865" height="70">
      <param name="movie" value="Flash/rodape_bayer.swf">
      <param name="quality" value="high">
      <embed src="Flash/rodape_bayer.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="865" height="70"></embed>
    </object></TD>
  </TR>  <tr>
    <td align="center" bgcolor="#000000"><p><font color="#FFFFFF" size="2" face="arial">Banda </font><font size="2" face="arial"><span class="style40">VP2</span></font><font color="#FFFFFF" size="2" face="arial"> - copyright 2009 - Todos os direitos reservados</font><br>
    </p>
    </td>
  </tr>
</table>
<?
 	}
?>
</body>
</html>