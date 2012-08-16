<%
'###################################################################
'##
'##  Script: ASPInfo v1.0
'##  Author: Michael Reisinger (OneWayMule)
'## 
'##  File: aspinfo.asp
'##
'##  
'##  This program is free software; you can redistribute it and/or
'##  modify it under the terms of the GNU General Public License
'##  as published by the Free Software Foundation. 
'## 
'##  This program is distributed in the hope that it will be useful,
'##  but WITHOUT ANY WARRANTY; without even the implied warranty of
'##  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'##  GNU General Public License for more details.
'##
'##  
'##  Support can be obtained from support forums at:
'##  http://www.onewaymule.org/onewayscripts/forums/
'## 
'##  You can get the latest version of this script at:
'##  http://www.onewaymule.org/onewayscripts/scripts/aspinfo/
'## 
'##  I greatly appreciate any comments, suggestions, bug reports. 
'##  Either post them at the forum (see above) or contact me at
'##  http://www.onewaymule.org/onewayscripts/contact.asp
'##
'###################################################################

Dim STR_COMPONENTS, ARR_COMPONENTS, OBJ_TEST, STR_KEY, INT_COUNT

STR_COMPONENTS = "" & _
"a1asp.crypto" & vbNewLine & _
"a1asp.dns" & vbNewLine & _
"a1asp.env" & vbNewLine & _
"a1asp.expand" & vbNewLine & _
"a1asp.fileop" & vbNewLine & _
"a1asp.fin" & vbNewLine & _
"a1asp.ini" & vbNewLine & _
"a1asp.reg" & vbNewLine & _
"a1asp.settings" & vbNewLine & _
"a1asp.sqs" & vbNewLine & _
"a1asp.syscontrol" & vbNewLine & _
"a1asp.sysinfo" & vbNewLine & _
"a1asp.uuidgen" & vbNewLine & _
"a1asp.wav" & vbNewLine & _
"ABCDrawHTML.Page" & vbNewLine & _
"ABCpdf3.Doc" & vbNewLine & _
"ABCpdf4.Doc" & vbNewLine & _
"ABCUpload4.XForm" & vbNewLine & _
"ABMailer.Mailman" & vbNewLine & _
"AbsoluteHttp.Conn" & vbNewLine & _
"AccessControl.SessAcc" & vbNewLine & _
"acDesktop.Desktop" & vbNewLine & _
"ACI.WhoIs" & vbNewLine & _
"acNetwork.DNS" & vbNewLine & _
"acSMTP.Smtp" & vbNewLine & _
"ACTarget.IPGEO" & vbNewLine & _
"ActiveFile.Archive" & vbNewLine & _
"Activefile.Directory" & vbNewLine & _
"ActiveFile.Post" & vbNewLine & _
"ActiveLaunch.Control" & vbNewLine & _
"ActiveMessenger.Message" & vbNewLine & _
"ActiveNavigator.Toolbar" & vbNewLine & _
"ActiveProfile.Profile" & vbNewLine & _
"ActiveSAR.SearchAndReplace" & vbNewLine & _
"ActiveShopper.BasketItem" & vbNewLine & _
"ActiveShopper.Cart" & vbNewLine & _
"Address2000.A2000" & vbNewLine & _
"AddressTools.EmailCheck" & vbNewLine & _
"AddressTools.ZIPCheck" & vbNewLine & _
"Adersoft.ImageGenerator" & vbNewLine & _
"ADISCON.SimpleMail" & vbNewLine & _
"ADISCON.SimpleMail.1" & vbNewLine & _
"AdminimizerLite.Editor" & vbNewLine & _
"ADODB.Command" & vbNewLine & _
"ADODB.Connection" & vbNewLine & _	
"ADODB.Recordset" & vbNewLine & _
"ADODB.Recordset.2.0" & vbNewLine & _
"ADODB.Recordset.2.7" & vbNewLine & _
"ADODB.Stream" & vbNewLine & _
"ADOX.Catalog" & vbNewLine & _
"AdvRegistry.Registry" & vbNewLine & _
"AesInterop.AesInterop" & vbNewLine & _
"Agni.PostToForm" & vbNewLine & _
"AlphaSierraPapa.AspRegSvr" & vbNewLine & _
"AMHTML.Application" & vbNewLine & _
"ANPOP.POPMAIN" & vbNewLine & _
"ANPOP.POPMSG" & vbNewLine & _
"ANSMTP.OBJ" & vbNewLine & _
"ANUPLOAD.OBJ" & vbNewLine & _
"APDocConv.Object" & vbNewLine & _
"ApolloASP60.Query" & vbNewLine & _
"ApolloASP61.Query" & vbNewLine & _
"APServer.Object" & vbNewLine & _
"APSpool.Object" & vbNewLine & _
"APToolkit.Object" & vbNewLine & _
"APWebGrabber.Object" & vbNewLine & _
"ASP.DB" & vbNewLine & _
"ASPBarChart100d.chart" & vbNewLine & _
"ASPCharge.CC" & vbNewLine & _
"ASPChart.Chart" & vbNewLine & _
"ASPControlHost.Host" & vbNewLine & _
"AspConv.Expert" & vbNewLine & _
"AspCrypt.Crypt" & vbNewLine & _
"aspCrypt.EasyCRYPT" & vbNewLine & _
"ASPCryptToy.CryptToy" & vbNewLine & _
"ASPdb.EP" & vbNewLine & _
"ASPdb.Free" & vbNewLine & _
"ASPdb.Pro" & vbNewLine & _
"ASPdb.View" & vbNewLine & _
"AspDbg.Trace" & vbNewLine & _
"ASPDNS.DNSLookup" & vbNewLine & _	
"AspDNS.Lookup" & vbNewLine & _
"ASPDns.ManageServer" & vbNewLine & _
"ASPExec.Execute" & vbNewLine & _
"AspFile.FileObj" & vbNewLine & _
"ASPFiles.Files" & vbNewLine & _
"ASPFileUpload.File" & vbNewLine & _
"AspHTTP.Conn" & vbNewLine & _
"AspImage.Image" & vbNewLine & _
"AspInet.FTP" & vbNewLine & _
"ASPL.Login" & vbNewLine & _
"ASPMAIL.ASPMailCtrl.1" & vbNewLine & _
"AspMap.Map" & vbNewLine & _
"AspMX.Lookup" & vbNewLine & _
"ASPMX.Resolver" & vbNewLine & _
"AspNNTP.Conn" & vbNewLine & _
"AspPager.Pager" & vbNewLine & _
"aspPDF.EasyPDF" & vbNewLine & _
"ASPPicture.Picture" & vbNewLine & _	
"AspPing.Conn" & vbNewLine & _
"ASPPlus.datacache" & vbNewLine & _
"aspPrint.ServerPrint" & vbNewLine & _
"AspQMail.Mailer" & vbNewLine & _
"ASPSimpleUpload.Upload" & vbNewLine & _
"AspSmartCache.SmartCache" & vbNewLine & _
"AspSmartChat.SmartChat" & vbNewLine & _
"AspSmartDate.SmartDate" & vbNewLine & _
"AspSmartFile.SmartFile" & vbNewLine & _
"AspSmartforum.smartforum" & vbNewLine & _
"AspSmartImage.SmartImage" & vbNewLine & _
"AspSmartMail.SmartMail" & vbNewLine & _
"AspSmartMenu.SmartMenuPopUp" & vbNewLine & _
"AspSmartUpload.SmartUpload" & vbNewLine & _
"ASPSMS.Booster" & vbNewLine & _	
"AspSock.Conn" & vbNewLine & _
"ASPsvg.Execute" & vbNewLine & _
"ASPsvg.Process" & vbNewLine & _
"ASPTabToy.TabToy" & vbNewLine & _
"ASPThumbnailer.Thumbnail" & vbNewLine & _
"ASPToday.VisitorID" & vbNewLine & _
"AspTouch.TouchIt" & vbNewLine & _
"AspUtils.TErrorManager" & vbNewLine & _
"AspUtils.TFileManager" & vbNewLine & _
"AspUtils.TFormManager" & vbNewLine & _
"AspUtils.tsmtpmanager" & vbNewLine & _
"AspUtils.TSQLManager" & vbNewLine & _
"AspWebCal120d.webcal" & vbNewLine & _
"ASPWordToy.WordToy" & vbNewLine & _
"ASPXP.Mail" & vbNewLine & _
"ASPXP2.WebPage" & vbNewLine & _
"aspZip.EasyZIP" & vbNewLine & _
"aspZipCodeToy.ZipCodeToy" & vbNewLine & _
"Atrax.ComboBox" & vbNewLine & _
"Atrax.EventLogger" & vbNewLine & _
"Atrax.Richedit" & vbNewLine & _
"Atrax.URLGrabber" & vbNewLine & _
"Atrax.Whois" & vbNewLine & _
"AuthNetSSLConnect.SSLPost" & vbNewLine & _
"AUTHXISP.AuthXOCXCtrl.1" & vbNewLine & _
"AUTHXOCX.AuthXOCXCtrl.1" & vbNewLine & _
"AXHostLib.Host" & vbNewLine & _
"AXSUPPORT.AXSupportCtrl.1" & vbNewLine & _
"Bamboo.SMTP" & vbNewLine & _
"basp21.FTP" & vbNewLine & _
"Bible.Lookup" & vbNewLine & _
"Bigtime.CSV" & vbNewLine & _
"BinaryFileStream.Object" & vbNewLine & _
"BinarySendFile.BinFileSend" & vbNewLine & _
"BitmapLibrary.NewIdentifier" & vbNewLine & _
"bkupload.form" & vbNewLine & _
"briz.AspThumb" & vbNewLine & _
"build.Report" & vbNewLine & _
"C2G.HTTP" & vbNewLine & _
"C2G.PING" & vbNewLine & _
"C2G.SCAN" & vbNewLine & _
"C2G.SCM" & vbNewLine & _
"C2G.TRACERT" & vbNewLine & _
"C2G.WHOIS" & vbNewLine & _
"c2geread.Message" & vbNewLine & _
"C2GSCM.Service" & vbNewLine & _
"CalendarCom.CalendarStuff" & vbNewLine & _
"CDO.Configuration" & vbNewLine & _
"CDO.MESSAGE" & vbNewLine & _
"CDONTS.NewMail" & vbNewLine & _
"CDOSYS.Message" & vbNewLine & _
"CFDEV.Activedit" & vbNewLine & _
"ChartDirector.API" & vbNewLine & _
"ChartFX.WebServer" & vbNewLine & _
"checkemail.maccheckemail" & vbNewLine & _
"ChilkatFTP.ChilkatFTP" & vbNewLine & _
"ChilkatImap.ChilkatImap" & vbNewLine & _
"ChilkatWebMail2.WebEmail2" & vbNewLine & _
"ChilkatXML.ChilkatXML" & vbNewLine & _
"ChilkatZip.ChilkatZip" & vbNewLine & _
"CMailCOM.SMTP.1" & vbNewLine & _
"cnhweb.http" & vbNewLine & _
"cnhweb.http2" & vbNewLine & _
"cnhweb.zip" & vbNewLine & _
"Coalesys.CSHttpClient.1" & vbNewLine & _
"Coalesys.CSPanelBar.2" & vbNewLine & _
"Coalesys.CSWebMenu.1" & vbNewLine & _
"com.comsoltech" & vbNewLine & _
"Commerce.DBStorage" & vbNewLine & _
"Commerce.Dictionary" & vbNewLine & _
"Commerce.ExpressionEvaluator" & vbNewLine & _
"commerce.micropipe" & vbNewLine & _
"commerce.mtspipeline" & vbNewLine & _
"commerce.orderform" & vbNewLine & _
"commerce.orderpipeline" & vbNewLine & _
"COMobjects.NET.PictureProcessor" & vbNewLine & _
"COMobjectsNET.Colorizer" & vbNewLine & _
"COMobjectsNET.IconGrabber" & vbNewLine & _
"COMobjectsNET.PictureGalleryPro" & vbNewLine & _
"COMobjectsNET.PieChart" & vbNewLine & _
"comUnixCrypt.UnixCryptObj" & vbNewLine & _
"Convert.t2h" & vbNewLine & _
"Corda.Embedder" & vbNewLine & _
"crossoft.quickcal" & vbNewLine & _
"crossoft.quicklist" & vbNewLine & _
"crossoft.quicktable" & vbNewLine & _
"crossoft.remotescript" & vbNewLine & _
"crossoft.waplist" & vbNewLine & _
"crossoft.wapsplash" & vbNewLine & _
"Crystal.CRPE.Application" & vbNewLine & _
"CrystalRuntime.application" & vbNewLine & _
"csASPUpload.Process" & vbNewLine & _
"csASPZipFile.MakeZip" & vbNewLine & _
"csImageFile.Manage" & vbNewLine & _
"csImageFileTrial.Manage" & vbNewLine & _
"CyberCashMCK.MessageBlock" & vbNewLine & _
"cyScape.browserObj" & vbNewLine & _
"cyScape.countryObj" & vbNewLine & _
"dafinfo.dafinfo.1" & vbNewLine & _
"DAO.DBEngine.35" & vbNewLine & _
"Dart.Dns" & vbNewLine & _
"Dart.Ftp" & vbNewLine & _
"Dart.Ftp.1" & vbNewLine & _
"Dart.Http" & vbNewLine & _
"Dart.Http.1" & vbNewLine & _
"Dart.Manager" & vbNewLine & _
"Dart.Manager.1" & vbNewLine & _
"Dart.Message" & vbNewLine & _
"Dart.Message.1" & vbNewLine & _
"Dart.Ping" & vbNewLine & _
"Dart.Ping.1" & vbNewLine & _
"Dart.Pop" & vbNewLine & _
"Dart.Pop.1" & vbNewLine & _
"Dart.Smtp" & vbNewLine & _
"Dart.Smtp.1" & vbNewLine & _
"Dart.Tcp" & vbNewLine & _
"Dart.Tcp.1" & vbNewLine & _
"Dart.Telnet" & vbNewLine & _
"Dart.Telnet.1" & vbNewLine & _
"Dart.WebASP" & vbNewLine & _
"Dart.WebASP.1" & vbNewLine & _
"Dart.WebPage" & vbNewLine & _
"Dart.WebPage.1" & vbNewLine & _
"Dart.Zip.1" & vbNewLine & _
"DartZip.Zip" & vbNewLine & _
"DartZip.Zip.1" & vbNewLine & _
"Datafun.FormBoy" & vbNewLine & _
"DbDataPrepComponent.DataPrep" & vbNewLine & _
"deruntime.deruntime" & vbNewLine & _
"desWIN.SysControl" & vbNewLine & _
"desWIN.SystemInfo" & vbNewLine & _	
"dgEncrypt.Key" & vbNewLine & _
"dgFileUpload.dgUpload" & vbNewLine & _
"dgReport.Report" & vbNewLine & _
"dgSort.QuickSort" & vbNewLine & _
"dgTree.Tree" & vbNewLine & _
"DigitalWol.Wol" & vbNewLine & _
"dImage.dThumb" & vbNewLine & _
"Dir2HTML.clsDirTree" & vbNewLine & _
"diwhois.diwhois" & vbNewLine & _
"dkqmail.qmail" & vbNewLine & _
"DrWFM.fm" & vbNewLine & _
"DTS.Package" & vbNewLine & _
"DTS.Packages" & vbNewLine & _
"Dundas.FilterPlus.1" & vbNewLine & _
"Dundas.Mailer" & vbNewLine & _
"Dundas.PieChartServer" & vbNewLine & _
"Dundas.PieChartServer.2" & vbNewLine & _
"Dundas.Upload" & vbNewLine & _
"Dundas.Upload.2" & vbNewLine & _
"Dundas.UploadProgress" & vbNewLine & _
"DynCrypto.Crypto" & vbNewLine & _
"DynImage.DynImage" & vbNewLine & _
"Dynu.CreditCard" & vbNewLine & _
"Dynu.DateTime" & vbNewLine & _
"Dynu.DNS" & vbNewLine & _
"Dynu.Email" & vbNewLine & _
"Dynu.Encrypt" & vbNewLine & _
"Dynu.Exec" & vbNewLine & _
"Dynu.FileUtil" & vbNewLine & _
"Dynu.FTP" & vbNewLine & _
"Dynu.GlobalWhois" & vbNewLine & _
"Dynu.HTTP" & vbNewLine & _
"Dynu.Ping" & vbNewLine & _
"Dynu.POP3" & vbNewLine & _
"Dynu.StringUtil" & vbNewLine & _
"Dynu.TCPSocket" & vbNewLine & _
"Dynu.Upload" & vbNewLine & _
"Dynu.Wait" & vbNewLine & _
"Dynu.Whois" & vbNewLine & _
"easyBarCode.aspBarCode" & vbNewLine & _
"EasyDb.Database" & vbNewLine & _
"EasyMail.SMTP.5" & vbNewLine & _
"EasyMail.SMTP.6" & vbNewLine & _
"ECHOCom.Echo" & vbNewLine & _
"EJB.stub" & vbNewLine & _
"ekov.PictureEffector" & vbNewLine & _
"eKov.PicturePreviewer" & vbNewLine & _
"eMarkASF.Movie" & vbNewLine & _
"emarkasi.painter" & vbNewLine & _
"Enom.EnomURL" & vbNewLine & _
"Excel.Application" & vbNewLine & _
"eyeXMail.CtrlClnt" & vbNewLine & _
"EZsite.CalendarManager" & vbNewLine & _
"EZsite.Calender" & vbNewLine & _
"EZsite.EZuploadLite" & vbNewLine & _
"EZsite.WebNotes" & vbNewLine & _
"FathMail.POP3" & vbNewLine & _
"FCCOM.ChartSrv" & vbNewLine & _
"FdfApp.FdfApp" & vbNewLine & _
"fdgfdgdfgdfgdf" & vbNewLine & _
"FileDownload.Manager" & vbNewLine & _
"FileSystem.FileSystemOb.1" & vbNewLine & _
"FileTouch.Control" & vbNewLine & _
"FreeINI.INI" & vbNewLine & _
"FTPort.FtpScript.1" & vbNewLine & _
"FuTime.UpTime" & vbNewLine & _
"GanX.AspFileUploader.1" & vbNewLine & _
"Geocel.Mailer" & vbNewLine & _
"GeoIPCOM.GeoIP" & vbNewLine & _
"GetServer.Status" & vbNewLine & _
"GFlax.GFlax" & vbNewLine & _
"GflAx170.GflAx" & vbNewLine & _
"GflAx193.GflAx" & vbNewLine & _
"GG_Perf.GetData" & vbNewLine & _
"GrabStock.GetQuote" & vbNewLine & _
"GraphicsProcessor2002.RasterObject" & vbNewLine & _
"grapl.engine" & vbNewLine & _
"GSDSvr.GSServerProp" & vbNewLine & _
"GSFile.Text" & vbNewLine & _
"GSGraph.Bar" & vbNewLine & _
"GSMail.Smtp" & vbNewLine & _
"GSServer.GSServerProp" & vbNewLine & _
"gsservices.settings" & vbNewLine & _
"GSSocket.TCPSock" & vbNewLine & _
"GSWhois.Whois" & vbNewLine & _
"GuidMakr.GUID" & vbNewLine & _
"HDSECompression.ActiveXZip.1" & vbNewLine & _
"HexDns.Connection" & vbNewLine & _
"Hexillion.HexIcmp" & vbNewLine & _
"Hexillion.HexLookup" & vbNewLine & _
"Hexillion.HexTcpQuery" & vbNewLine & _
"HexValidEmail.Connection" & vbNewLine & _
"Hommingberger.Gepardenforelle" & vbNewLine & _
"Htmlizer.Text" & vbNewLine & _
"HTTPheadInfo.DisplayHTTP" & vbNewLine & _
"id3.id3get" & vbNewLine & _
"IDSMailInterface.Server" & vbNewLine & _
"IImageGlue5.Graphic" & vbNewLine & _
"iiscart2000.store" & vbNewLine & _
"iisCC.cc" & vbNewLine & _
"IISmail.iismail" & vbNewLine & _
"IISmail.iismail.1" & vbNewLine & _
"IISSample.asp2htm" & vbNewLine & _
"IISSample.contentrotator" & vbNewLine & _
"IISSample.LookupTable" & vbNewLine & _
"IISSample.registry" & vbNewLine & _
"IISSample.summaryinfos" & vbNewLine & _
"IISSample.tracer" & vbNewLine & _
"ImageEnASP.ImageEn" & vbNewLine & _
"ImageGlue.Canvas" & vbNewLine & _
"ImageGlue5.Canvas" & vbNewLine & _
"ImageGlue5.Gestalt" & vbNewLine & _
"ImageGlue5.XRect" & vbNewLine & _
"ImageGoo.XMultipart" & vbNewLine & _
"ImgSize.Check" & vbNewLine & _
"ImgXASP6.ImgX" & vbNewLine & _
"InetCtls.Inet" & vbNewLine & _
"InetCtls.Inet.1" & vbNewLine & _
"InteliSource.Online" & vbNewLine & _
"IntrCard.Credit" & vbNewLine & _
"IntrChart.Chart" & vbNewLine & _
"IntrPWD.Validate" & vbNewLine & _
"IntrSQL.Query" & vbNewLine & _
"iPlotLibrary.iPlotX" & vbNewLine & _
"IPWorksASP.FileMailer" & vbNewLine & _
"IPWorksASP.FTP" & vbNewLine & _
"IPWorksASP.HTMLMailer" & vbNewLine & _
"IPWorksASP.HTTP" & vbNewLine & _
"IPWorksASP.ICMPPort" & vbNewLine & _
"IPWorksASP.IMAP" & vbNewLine & _
"IPWorksASP.IPInfo" & vbNewLine & _
"IPWorksASP.IPPort" & vbNewLine & _
"IPWorksASP.LDAP" & vbNewLine & _
"IPWorksASP.MCast" & vbNewLine & _
"IPWorksASP.MIME" & vbNewLine & _
"IPWorksASP.MX" & vbNewLine & _
"IPWorksASP.NetClock" & vbNewLine & _
"IPWorksASP.NetCode" & vbNewLine & _
"IPWorksASP.NetDial" & vbNewLine & _
"IPWorksASP.NNTP" & vbNewLine & _
"IPWorksASP.Ping" & vbNewLine & _
"IPWorksASP.POP" & vbNewLine & _
"IPWorksASP.RCP" & vbNewLine & _
"IPWorksASP.Rexec" & vbNewLine & _
"IPWorksASP.Rshell" & vbNewLine & _
"IPWorksASP.SMTP" & vbNewLine & _
"IPWorksASP.SNMP" & vbNewLine & _
"IPWorksASP.SNPP" & vbNewLine & _
"IPWorksASP.SOAP" & vbNewLine & _
"IPWorksASP.Telnet" & vbNewLine & _
"IPWorksASP.TFTP" & vbNewLine & _
"IPWorksASP.TraceRoute" & vbNewLine & _
"IPWorksASP.UDPPort" & vbNewLine & _
"IPWorksASP.WebForm" & vbNewLine & _
"IPWorksASP.WebUpload" & vbNewLine & _
"IPWorksASP.Whois" & vbNewLine & _
"IPWorksASP.XMLp" & vbNewLine & _
"ISSecureFile.FileSystemObject" & vbNewLine & _
"ITDNpxnbusiness.pxnbusiness" & vbNewLine & _
"ixsso.Query" & vbNewLine & _
"ixsso.Util" & vbNewLine & _
"Jaguar.ORB" & vbNewLine & _
"JAL.SQL2Table" & vbNewLine & _
"JALAppt.Appointment" & vbNewLine & _
"JALCal.Calendar" & vbNewLine & _
"janGraphics.Compendium" & vbNewLine & _
"JavaPop3.Mailer" & vbNewLine & _
"javaside.rbl.acximage" & vbNewLine & _
"JESoftware.xContent" & vbNewLine & _
"JESoftware.xPop3" & vbNewLine & _
"JESoftware.xSMTP" & vbNewLine & _
"JESoftware.xTree" & vbNewLine & _
"JMail.Message" & vbNewLine & _
"JMail.POP3" & vbNewLine & _
"JMail.SMTPMail" & vbNewLine & _
"JpegLib.cJpeg" & vbNewLine & _
"JRO.JetEngine" & vbNewLine & _
"JustLDAP.Find" & vbNewLine & _
"khttp.inet" & vbNewLine & _
"KISSW3.Application" & vbNewLine & _
"LastMod.FileObj" & vbNewLine & _
"LENNE.Compactor" & vbNewLine & _
"lightcom.xBrowser" & vbNewLine & _
"lightcom.xContent" & vbNewLine & _
"lightcom.xPop3" & vbNewLine & _
"lightcom.xSMTP" & vbNewLine & _
"lightcom.xTree" & vbNewLine & _
"LiveSoup.OpenSMSLite" & vbNewLine & _
"LookupTable.cLookupTable" & vbNewLine & _
"lo_Login.clsLogin" & vbNewLine & _
"LpiCom_5_2.LinkPointCom" & vbNewLine & _
"LpiCom_5_4.LinkPointCom" & vbNewLine & _
"LpiCom_6_0.LPOrderPart" & vbNewLine & _
"lyfimage.image" & vbNewLine & _
"LyfUpload.UploadFile" & vbNewLine & _
"Lyris.LCP" & vbNewLine & _
"MagicBundle.MagicINI" & vbNewLine & _
"MagicBundle.MagicRegistry" & vbNewLine & _
"MagicRegistry.Tricks" & vbNewLine & _
"MailBee.POP3" & vbNewLine & _
"MailBee.SMTP" & vbNewLine & _
"MailServerX.Users" & vbNewLine & _
"Majodio.FTP" & vbNewLine & _
"MAPI.Session" & vbNewLine & _
"MarkItUp.Configuration" & vbNewLine & _
"MD5DLL.Encrypt" & vbNewLine & _
"MEMail.Message" & vbNewLine & _
"MfsManage.MfsSession" & vbNewLine & _
"Microsoft.DiskQuota.1" & vbNewLine & _
"Microsoft.XMLDOM" & vbNewLine & _
"Microsoft.XMLHTTP" & vbNewLine & _
"Miraplacid.MSCCryptoAES" & vbNewLine & _
"Miraplacid.MSCExec" & vbNewLine & _
"MJRSChat.WebChat" & vbNewLine & _
"MP.Filtering" & vbNewLine & _
"MPS.SendMail" & vbNewLine & _
"MP_Mikys_ASP.Password" & vbNewLine & _
"MS.Finance" & vbNewLine & _
"MSCommLib.MSComm" & vbNewLine & _
"MSCommLIB.MSComm.1" & vbNewLine & _
"MSMQ.MSMQQueue" & vbNewLine & _
"MSMQ.MSMQQueueInfo" & vbNewLine & _
"mssoap.soapclient" & vbNewLine & _
"MSSOAP.SoapClient.1" & vbNewLine & _
"MSSOAP.SOAPClient30" & vbNewLine & _
"MSWC.AdRotator" & vbNewLine & _
"MSWC.BrowserType" & vbNewLine & _
"MSWC.ContentRotator" & vbNewLine & _
"MSWC.Counters" & vbNewLine & _
"MSWC.IISLog" & vbNewLine & _
"MSWC.loadbalance" & vbNewLine & _
"MSWC.NextLink" & vbNewLine & _
"MSWC.PageCounter" & vbNewLine & _
"MSWC.PermissionChecker" & vbNewLine & _
"MSWC.Status" & vbNewLine & _
"MSWC.Tools" & vbNewLine & _
"MSWinsock.Winsock" & vbNewLine & _
"MSXML.ServerXMLHTTP" & vbNewLine & _
"MSXML2.DOMDocument" & vbNewLine & _
"MSXML2.DOMDocument.2.6" & vbNewLine & _	
"MSXML2.DOMDocument.3.0" & vbNewLine & _
"MSXML2.DOMDocument.4.0" & vbNewLine & _
"MSXML2.DOMDocument.5.0" & vbNewLine & _
"MSXML2.FreeThreadedDOMDocument.3.0" & vbNewLine & _
"MSXML2.FreeThreadedDOMDocument.4.0" & vbNewLine & _
"MSXML2.ServerXMLHTTP" & vbNewLine & _
"MSXML2.ServerXMLHTTP.3.0" & vbNewLine & _
"MSXML2.ServerXMLHTTP.4.0" & vbNewLine & _
"MSXML2.XMLSchemaCache.4.0" & vbNewLine & _
"MSXML2.XSLTemplate" & vbNewLine & _
"MSXML2.XSLTemplate.4.0" & vbNewLine & _
"MSXML3.DomDocument" & vbNewLine & _
"MSXML3.ServerXMLHTTP" & vbNewLine & _
"MSXML4.DomDocument" & vbNewLine & _
"MSXML4.ServerXMLHTTP" & vbNewLine & _
"Ncdo.Ncdonts.1" & vbNewLine & _
"NCWebToy.ASPHTM" & vbNewLine & _
"NERV.Network" & vbNewLine & _
"NetCAuth.NetCObj.1" & vbNewLine & _
"NetChartsServer.NSToolKit" & vbNewLine & _
"NETDLL.Network" & vbNewLine & _
"NETOMATIX.ImageEngine" & vbNewLine & _
"NETOMATIX.ImageServer" & vbNewLine & _
"newObjects.utilctls.SFMain" & vbNewLine & _
"NisCom.ImageComponent" & vbNewLine & _
"NWDirLib.NWDirCtrl.1" & vbNewLine & _
"objBarGraph.DrawChart" & vbNewLine & _
"obout_ASPTreeView.Tree" & vbNewLine & _
"obout_ASPTreeView_Pro.Tree" & vbNewLine & _
"obout_ASPTreeView_XP.Tree" & vbNewLine & _
"obout_SlideMenu_Pro.MenuPro" & vbNewLine & _
"OCXHTTP.OCXHttpCtrl" & vbNewLine & _
"OCXHTTP.OCXHttpCtrl.1" & vbNewLine & _
"OCXQmail.OCXQmailCtrl" & vbNewLine & _
"OCXQmail.OCXQmailCtrl.1" & vbNewLine & _
"OneTouchASP.StrFunctions" & vbNewLine & _
"OpenX.DBMail" & vbNewLine & _
"OpenX2.Connection" & vbNewLine & _
"OracleInProcServer.XOraSession" & vbNewLine & _
"OrgChart.Tree" & vbNewLine & _
"OSSMTP.SMTPSession" & vbNewLine & _
"Outlook.Application" & vbNewLine & _
"Overpower.ImageLib" & vbNewLine & _
"OWC.Chart" & vbNewLine & _
"OWC.spreadsheet" & vbNewLine & _
"OWC.spreadsheet.9" & vbNewLine & _
"OWC10.ChartSpace" & vbNewLine & _
"PCAuthX.Authorizer" & vbNewLine & _
"PDIRTFCONVERTER.DocConverter" & vbNewLine & _
"Persits.Aspemail" & vbNewLine & _
"Persits.Aspencrypt" & vbNewLine & _
"Persits.AspUser" & vbNewLine & _
"Persits.CryptoManager" & vbNewLine & _
"Persits.Grid" & vbNewLine & _
"Persits.Jpeg" & vbNewLine & _
"Persits.MailSender" & vbNewLine & _
"persits.pdf" & vbNewLine & _
"Persits.Upload" & vbNewLine & _
"Persits.Upload.1" & vbNewLine & _
"persits.upload.3" & vbNewLine & _
"Persits.UploadProgress" & vbNewLine & _
"Persits.XUpload" & vbNewLine & _
"PFProCOMControl.PFProCOMControl.1" & vbNewLine & _
"Photoshop.ActionDescriptor" & vbNewLine & _
"Photoshop.Application" & vbNewLine & _
"Pnvzip.ZipFunctions" & vbNewLine & _
"POP3svg.Mailer" & vbNewLine & _
"port80.httpzip" & vbNewLine & _
"ppthumbnail.cthumbgen" & vbNewLine & _
"ProjectDisplay.Charts" & vbNewLine & _
"Prt2Disk.Control" & vbNewLine & _
"quicktab.quicktabs" & vbNewLine & _
"QwerkSoft.FormSlam" & vbNewLine & _
"RBarcode.RBarcodeX" & vbNewLine & _
"RC4DLL.Crypt" & vbNewLine & _
"ReportMan.ReportManX" & vbNewLine & _
"RSADLL.Crypt" & vbNewLine & _
"RSADLL1.KeyGen" & vbNewLine & _
"runwaycore.base64" & vbNewLine & _
"runwaycore.builder" & vbNewLine & _
"runwaycore.data" & vbNewLine & _
"runwaycore.files" & vbNewLine & _
"runwaycore.ftp" & vbNewLine & _
"runwaycore.md5" & vbNewLine & _
"runwaycore.rc4" & vbNewLine & _
"runwaycore.smtp" & vbNewLine & _
"runwaycore.tcp" & vbNewLine & _
"S3Weather.Current" & vbNewLine & _
"SAPI.SpVoice" & vbNewLine & _
"SASWorkspaceManager.WorkspaceManager" & vbNewLine & _
"Schroeder.Traffictool" & vbNewLine & _
"SciBit.MySQLX" & vbNewLine & _
"SCP.AdminUtils" & vbNewLine & _
"Scribe.ScribeDOM" & vbNewLine & _
"Scripting.Dictionary" & vbNewLine & _
"Scripting.FileSystemObject" & vbNewLine & _
"Scriptlet.TypeLib" & vbNewLine & _
"ScriptUtils.ASPForm" & vbNewLine & _
"ScriptUtils.ByteArray" & vbNewLine & _
"ScriptUtils.Kernel" & vbNewLine & _
"Search.SearchAdmin.1" & vbNewLine & _
"SearchWorks.FileSearch" & vbNewLine & _
"SemClient.Control" & vbNewLine & _
"Session.Management" & vbNewLine & _
"SfImageResize.ImageResize" & vbNewLine & _
"shotgraph.image" & vbNewLine & _
"SHOTIP.Connection" & vbNewLine & _
"SImageUtil.Image" & vbNewLine & _
"SimplePageASP.SNPP" & vbNewLine & _
"Simplewire.SMS" & vbNewLine & _
"SimplyMailer.SMTP" & vbNewLine & _
"SiteAdmin.AdminTools" & vbNewLine & _
"SiteSecurity.Login" & vbNewLine & _
"Sloppycode.DataProducer" & vbNewLine & _
"SMS22.SendSMS22" & vbNewLine & _
"smtp" & vbNewLine & _
"SmtpMail.SmtpMail" & vbNewLine & _
"SmtpMail.SmtpMail.1" & vbNewLine & _
"SMTPsvg.Mailer" & vbNewLine & _
"SMUM.XCheck" & vbNewLine & _
"SMUM.XCheck.1" & vbNewLine & _
"Socket.TCP" & vbNewLine & _
"Softartisans.Archive" & vbNewLine & _
"Softartisans.ExcelWriter" & vbNewLine & _
"SoftArtisans.FileManager" & vbNewLine & _
"SoftArtisans.FileManagerTX" & vbNewLine & _
"SoftArtisans.FileUp" & vbNewLine & _
"SoftArtisans.FileUpProgress" & vbNewLine & _
"SoftArtisans.Groups" & vbNewLine & _
"softartisans.ImageGen" & vbNewLine & _
"softartisans.jfile" & vbNewLine & _
"SoftArtisans.Performance" & vbNewLine & _
"SoftArtisans.RAS" & vbNewLine & _
"SoftArtisans.SACheck" & vbNewLine & _
"SoftArtisans.SAFile" & vbNewLine & _
"SoftArtisans.SASessionPro" & vbNewLine & _
"SoftArtisans.SASessionPro.1" & vbNewLine & _
"SoftArtisans.Shares" & vbNewLine & _
"SoftArtisans.SMTPMail" & vbNewLine & _
"SoftArtisans.User" & vbNewLine & _
"SoftArtisans.XFile" & vbNewLine & _
"SoftArtisans.XFRequest" & vbNewLine & _
"SoftComplex.ASP.PostStream" & vbNewLine & _
"SoftComplex.Email" & vbNewLine & _
"SoftComplex.Zip" & vbNewLine & _
"SOFTOMATIX.ChartEngine" & vbNewLine & _
"Softwing.ASPEventlog" & vbNewLine & _
"Softwing.AspQPerfCounters" & vbNewLine & _
"Softwing.AspTear" & vbNewLine & _
"Softwing.EDConverter" & vbNewLine & _
"Softwing.EventLogReader" & vbNewLine & _
"Softwing.FileCache.1" & vbNewLine & _
"Softwing.LocaleFormatter" & vbNewLine & _
"Softwing.MacBinary" & vbNewLine & _
"Softwing.OdbcRegTool" & vbNewLine & _
"Softwing.Profiler" & vbNewLine & _
"Softwing.VersionInfo" & vbNewLine & _
"SoftwingXSB.ShoppingBag" & vbNewLine & _
"SPCryptToy.CryptToy" & vbNewLine & _
"SPrinterPro.Object" & vbNewLine & _
"SQLDataSoap.SoapAgent" & vbNewLine & _
"SQLDMO.Database" & vbNewLine & _
"SQLDMO.SQLServer" & vbNewLine & _
"SQLOLE.SQLServer" & vbNewLine & _
"SQLsearch.read" & vbNewLine & _
"sqlxmlbulkload.SQLXMLBulkLoad" & vbNewLine & _
"Statron.Control" & vbNewLine & _
"Stonebroom.ASP2XML" & vbNewLine & _
"Stonebroom.ASPointer" & vbNewLine & _
"Stonebroom.RegEx" & vbNewLine & _
"Stonebroom.RemoteZip" & vbNewLine & _
"Stonebroom.SaveForm" & vbNewLine & _
"Stonebroom.ServerZip" & vbNewLine & _
"Stonebroom.XSLTransform" & vbNewLine & _
"swfobjs.swfMovie" & vbNewLine & _
"swfobjs.swfObject" & vbNewLine & _
"SX.Color_Converter" & vbNewLine & _
"SystemDaren.DarenSystemInfo" & vbNewLine & _
"T1CFreeImage.Images.1" & vbNewLine & _
"Tabliz.AdminRecordset" & vbNewLine & _
"TCPIP.DNS" & vbNewLine & _
"TCPIP.Trace" & vbNewLine & _
"TeamKeso.ULoad" & vbNewLine & _
"Temperature.Conversion" & vbNewLine & _
"TeriaWC.WebCalendar" & vbNewLine & _
"Text2Tree150d.tree" & vbNewLine & _
"TideServer.TideChart" & vbNewLine & _
"TimeSpan.Control" & vbNewLine & _
"tntoday.news" & vbNewLine & _
"TreeGen.Tree" & vbNewLine & _
"UDDI10.find_business" & vbNewLine & _
"UltraWOL.ctlUltraWOL" & vbNewLine & _
"UnitedBinary.AutoImageS" & vbNewLine & _
"UnitedBinary.AutoImageSize" & vbNewLine & _
"URLFetch.URLFetch" & vbNewLine & _
"UserManager.Server" & vbNewLine & _
"VASPLV.ASPListView" & vbNewLine & _
"VASPMV.ASPMonthView" & vbNewLine & _
"VASPTB.ASPTabView" & vbNewLine & _
"VASPTV.ASPTreeView" & vbNewLine & _
"VBScript.RegExp" & vbNewLine & _
"VIMAS.Image" & vbNewLine & _
"VisualSoft.BLOWFISHCrypt" & vbNewLine & _
"VisualSoft.BLOWFISHCrypt.1" & vbNewLine & _
"VisualSoft.Chart" & vbNewLine & _
"VisualSoft.Chart.1" & vbNewLine & _
"VisualSoft.DataAdmin" & vbNewLine & _
"VisualSoft.DataAdmin.1" & vbNewLine & _
"VisualSoft.DMXML" & vbNewLine & _
"VisualSoft.DMXML.1" & vbNewLine & _
"VisualSoft.FTP" & vbNewLine & _
"VisualSoft.FTP.1" & vbNewLine & _
"VisualSoft.HTTP" & vbNewLine & _
"VisualSoft.HTTP.1" & vbNewLine & _
"VisualSoft.Mail" & vbNewLine & _
"VisualSoft.Mail.1" & vbNewLine & _
"VoiceShot.VoiceShot" & vbNewLine & _
"w3.netutils" & vbNewLine & _
"w3.upload" & vbNewLine & _
"W3Image.Image" & vbNewLine & _
"w3info.w3info.1" & vbNewLine & _
"w3sitetree.tree" & vbNewLine & _
"WaitFor.Comp" & vbNewLine & _
"waspzip.waspzip" & vbNewLine & _
"WB.FadeBulletin" & vbNewLine & _
"WbemScripting.SWbemLocator" & vbNewLine & _
"wddx.deserializer.1" & vbNewLine & _
"WDDX.Serializer.1" & vbNewLine & _
"WDDX.Struct.1" & vbNewLine & _
"webbuilder.wbASPWebBuilder" & vbNewLine & _
"webcam32.application" & vbNewLine & _
"WebStormMail.POP3.1" & vbNewLine & _
"WebStormMail.SMTP.1" & vbNewLine & _
"werkslib.mp3exp" & vbNewLine & _
"WhoIs2.WhoIs" & vbNewLine & _
"WhoIs3.WhoIs" & vbNewLine & _
"WhoisDLL.Whois" & vbNewLine & _
"WinASP.FileAction" & vbNewLine & _
"Wingraphviz.dot" & vbNewLine & _
"WinHttp.WinHttpRequest.5" & vbNewLine & _
"WinHttp.WinHttpRequest.5.1" & vbNewLine & _
"word.application" & vbNewLine & _
"WorldPay.COMpurchase" & vbNewLine & _
"WorldTime.Engine" & vbNewLine & _
"WScript.Network" & vbNewLine & _
"WScript.Shell" & vbNewLine & _
"wshMathFunctions.ucMath" & vbNewLine & _
"wt_sendmail.send" & vbNewLine & _
"WWWPrint.Client" & vbNewLine & _
"xAddress.Process" & vbNewLine & _
"xAuthorize.Charge" & vbNewLine & _
"XceedSoftware.XceedFtp" & vbNewLine & _
"XceedSoftware.XceedZip.4" & vbNewLine & _
"xceedsoftware.xceedzip.5" & vbNewLine & _
"xcrypt.crypt" & vbNewLine & _
"xders.Recordset" & vbNewLine & _
"XMComCRC.XMCommCRC" & vbNewLine & _
"XMLRPC.Convert" & vbNewLine & _
"xShip.Rates" & vbNewLine & _
"XStandard.Base64" & vbNewLine & _
"XStandard.Buffer" & vbNewLine & _
"XStandard.CSS2XML" & vbNewLine & _
"XStandard.GZip" & vbNewLine & _
"XStandard.HTTP" & vbNewLine & _
"XStandard.Image" & vbNewLine & _
"XStandard.ISO8601" & vbNewLine & _
"XStandard.MD5" & vbNewLine & _
"XStandard.TAR" & vbNewLine & _
"XStandard.Zip" & vbNewLine & _
"YesFunc.Func" & vbNewLine & _
"ZaksPop3.Server" & vbNewLine & _
"zbitz.zzip" & vbNewLine & _
"Zikit.Hap" & vbNewLine & _
"zimage.ServerFonts" & vbNewLine & _
"zimage.zimage" & vbNewLine & _
"ZmeYsoft.Hashes.MD5" & vbNewLine & _
"ZmeYsoft.Util.ImageFile" & vbNewLine & _	
"ZonerWeb.Tools" & vbNewLine
ARR_COMPONENTS = Split(STR_COMPONENTS, vbNewline)


Sub Display_Header()
	Response.Write  "	<p>" & vbNewLine & _
			"	<table class=""main"" cellspacing=""1"" align=""center"" style=""width:90%;"">" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header"" align=""left"" colspan=""2"">&nbsp;<font size=""3"">ASPInfo</font>&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			Display_Data("Date", now()) & _
			Display_Data("Script&nbsp;URL", "<a href=""http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO") & """>http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO") & "</a>") & _
			Display_Data("Physical&nbsp;Path", Server.MapPath(Request.ServerVariables("SCRIPT_NAME"))) & _
			"	</table>" & vbNewLine & _
			"	</p>" & vbNewLine
End Sub


Sub Display_Server_Information()
	Response.Write  "	<p>" & vbNewLine & _
			"	<table class=""main"" cellspacing=""1"" align=""center"" style=""width:90%;"">" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header"" colspan=""2"" align=""center"">&nbsp;Server&nbsp;Information&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Info&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			Display_Data("Server&nbsp;Name", Request.ServerVariables("SERVER_NAME")) & _
			Display_Data("Server&nbsp;IP&nbsp;Address", Request.ServerVariables("LOCAL_ADDR")) & _
			Display_Data("Server&nbsp;Port", Request.ServerVariables("SERVER_PORT")) & _
			Display_Data("Server&nbsp;Software", Request.ServerVariables("SERVER_SOFTWARE")) & _
			Display_Data("Operating&nbsp;System", Request.ServerVariables("OS")) & _
			Display_Data("Script&nbsp;Engine", ScriptEngine & " (version: " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion & ")") & _
			"	</table>" & vbNewLine & _
			"	</p>" & vbNewLine
End Sub


Sub Display_Application_Object()
	Response.Write  "	<p>" & vbNewLine & _
			"	<table class=""main"" cellspacing=""1"" align=""center"" style=""width:90%;"">" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header"" colspan=""2"" align=""center"">&nbsp;Application Object&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Variable&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine
	INT_COUNT = 0
	For Each STR_KEY In Application.Contents

		If isObject(Application.Contents(STR_KEY)) Then
			Response.Write  Display_Data("Application(""" & STR_KEY & """)", "(Object)")
		Else
			Response.Write  Display_Data("Application(""" & STR_KEY & """)", Application.Contents(STR_KEY))
		End If
		Response.Write	"		</tr>" & vbNewLine
		INT_COUNT = INT_COUNT + 1
	Next
	If INT_COUNT = 0 Then
		Response.Write  "		<tr>" & vbNewLine & _
				"			<td class=""cell"" align=""left"" colspan=""2"">(no data)</td>" & vbNewLine & _
				"		</tr>"
	End If
	Response.Write  "	</table>" & vbNewLine & _
			"	</p>" & vbNewLine
End Sub


Sub Display_Request_Object()
	Response.Write  "	<p>" & vbNewLine & _
			"	<table class=""main"" cellspacing=""1"" align=""center"" style=""width:90%;"">" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header"" colspan=""2"" align=""center"">&nbsp;Request&nbsp;Object&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;ClientCertificate&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine
	INT_COUNT = 0
	For Each STR_KEY In Request.ClientCertificate
		Response.Write  Display_Data ("Request.ClientCertificate(""" & STR_KEY & """)", Request.ClientCertificate(STR_KEY))
		INT_COUNT = INT_COUNT + 1
	Next
	If INT_COUNT = 0 Then
		Response.Write  "		<tr>" & vbNewLine & _
				"			<td class=""cell"" align=""left"" colspan=""2"">(no data)</td>" & vbNewLine & _
				"		</tr>"
	End If
	Response.Write	"       	<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Cookies&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine
	INT_COUNT = 0
	For Each STR_KEY In Request.Cookies
		If Request.Cookies(STR_KEY).HasKeys Then
			For Each STR_KEY2 In Request.Cookies(STR_KEY)
				Response.Write  Display_Data("Cookies(""" & STR_KEY & """)(""" & STR_KEY2 & """)", Request.Cookies(STR_KEY)(STR_KEY2))
			Next
		Else
			Response.Write  Display_Data ("Request.Cookies(""" & STR_KEY & """)", Request.Cookies(STR_KEY))
		End If
		INT_COUNT = INT_COUNT + 1
	Next
	If INT_COUNT = 0 Then
		Response.Write  "		<tr>" & vbNewLine & _
				"			<td class=""cell"" align=""left"" colspan=""2"">(no data)</td>" & vbNewLine & _
				"		</tr>"
	End If
	Response.Write	"		<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Form&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine
	INT_COUNT = 0
	For Each STR_KEY In Request.Form
		Response.Write  Display_Data ("Request.Form(""" & STR_KEY & """)", Request.Form(STR_KEY))
		INT_COUNT = INT_COUNT + 1
	Next
	If INT_COUNT = 0 Then
		Response.Write  "		<tr>" & vbNewLine & _
				"			<td class=""cell"" align=""left"" colspan=""2"">(no data)</td>" & vbNewLine & _
				"		</tr>"
	End If
	Response.Write	"       	<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;QueryString&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine
	INT_COUNT = 0
	For Each STR_KEY In Request.QueryString
		Response.Write  Display_Data ("Request.QueryString(""" & STR_KEY & """)", QueryString(STR_KEY))
		INT_COUNT = INT_COUNT + 1
	Next
	If INT_COUNT = 0 Then
		Response.Write  "		<tr>" & vbNewLine & _
				"			<td class=""cell"" align=""left"" colspan=""2"">(no data)</td>" & vbNewLine & _
				"		</tr>"
	End If
	Response.Write	"       	<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;ServerVariables&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine
	INT_COUNT = 0
	For Each STR_KEY In Request.ServerVariables
		Response.Write  Display_Data ("Request.ServerVariables(""" & STR_KEY & """)", Request.ServerVariables(STR_KEY))
		INT_COUNT = INT_COUNT + 1
	Next
	If INT_COUNT = 0 Then
		Response.Write  "		<tr>" & vbNewLine & _
				"			<td class=""cell"" align=""left"" colspan=""2"">(no data)</td>" & vbNewLine & _
				"		</tr>"
	End If
	Response.Write  "	</table>" & vbNewLine & _
			"	</p>" & vbNewLine
End Sub


Sub Display_Response_Object()
	Response.Write  "	<p>" & vbNewLine & _
			"	<table class=""main"" cellspacing=""1"" align=""center"" style=""width:90%;"">" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header"" colspan=""2"" align=""center"">&nbsp;Response Object&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Attribute&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			Display_Data("Response.Buffer", Response.Buffer) & _
			Display_Data("Response.CacheControl", Response.CacheControl) & _
			Display_Data("Response.Charset", Response.Charset) & _
			Display_Data("Response.ContentType", Response.ContentType) & _
			Display_Data("Response.Expires", Response.Expires) & _
			Display_Data("Response.ExpiresAbsolute", Response.ExpiresAbsolute) & _
			Display_Data("Response.isClientConnected", Response.isClientConnected) & _
			Display_Data("Response.Status", Response.Status) & _
			"	</table>" & vbNewLine & _
			"	</p>" & vbNewLine
End Sub


Sub Display_Server_Object()
	Response.Write  "	<p>" & vbNewLine & _
			"	<table class=""main"" cellspacing=""1"" align=""center"" style=""width:90%;"">" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header"" colspan=""2"" align=""center"">&nbsp;Server&nbsp;Object&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Attribute&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			Display_Data("Server.ScriptTimeout", Server.ScriptTimeout & "&nbsp;</font>") & _ 
			"	</table>" & vbNewLine & _
			"	</p>" & vbNewLine
End Sub


Sub Display_Session_Object()
	Response.Write  "	<p>" & vbNewLine & _
			"	<table class=""main"" cellspacing=""1"" align=""center"" style=""width:90%;"">" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header"" colspan=""2"" align=""center"">&nbsp;Session Object&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Attribute&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			Display_Data("Session.CodePage", Session.CodePage) & _
			Display_Data("Session.LCID", Session.LCID) & _
			Display_Data("Session.SessionID", Session.SessionID) & _
			Display_Data("Session.TimeOut", Session.TimeOut) & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Variable&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine
	INT_COUNT = 0
	For Each STR_KEY In Session.Contents
		If isObject(Session.Contents("STR_KEY")) Then
			Response.Write Display_Data("Session(""" & STR_KEY & """)", "(Object)")			
		Else
			Response.Write Display_Data("Session(""" & STR_KEY & """)", Session("STR_KEY"))
		End If
		INT_COUNT = INT_COUNT + 1
	Next
	If INT_COUNT = 0 Then
		Response.Write  "		<tr>" & vbNewLine & _
				"			<td class=""cell"" align=""left"" colspan=""2"">(no data)</td>" & vbNewLine & _
				"		</tr>"
	End If
	Response.Write  "       	<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;StaticObject&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Value&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine
	INT_COUNT = 0
	For Each STR_KEY In Session.StaticObjects
		If isObject(Session.StaticObjects("STR_KEY")) Then
	        	Response.Write Display_Data("Session.StaticObjects(""" & STR_KEY & """)", "(Object)")
		Else
        		Response.Write Display_Data("Session.StaticObjects(""" & STR_KEY & """)", Session.StaticObjects(STR_KEY))
		End If
		INT_COUNT = INT_COUNT + 1
	Next
	If INT_COUNT = 0 Then
		Response.Write  "		<tr>" & vbNewLine & _
				"			<td class=""cell"" align=""left"" colspan=""2"">(no data)</td>" & vbNewLine & _
				"		</tr>"
	End If
	Response.Write  "	</table>" & vbNewLine & _
			"	</p>" & vbNewLine
End Sub


Sub Display_All_Components()
	Dim STR_TEMP, INT_COUNT2
	STR_TEMP = ""
	INT_COUNT = 0
	On Error Resume Next
	For INT_COUNT2 = 0 To UBound(ARR_COMPONENTS)
		Err = 0
		STR_COMPONENT = ARR_COMPONENTS(INT_COUNT2)
		STR_VER = ""
		Set OBJ_TEST = Server.CreateObject(STR_COMPONENT)
		If Err = 0 Then
			STR_VER = OBJ_TEST.Version
			STR_TEMP = STR_TEMP & Display_Data(STR_COMPONENT,"" & STR_VER & "")
			INT_COUNT = INT_COUNT + 1
		End If
		Set OBJ_TEST = Nothing
	Next
	On Error Goto 0
	STR_TEMP =      "	<p>" & vbNewLine & _
			"	<table class=""main"" cellspacing=""1"" align=""center"" style=""width:90%;"">" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header"" colspan=""2"" align=""center"">&nbsp;Installed&nbsp;Components&nbsp;(" & INT_COUNT & "&nbsp;installed)</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Component&nbsp;</td>" & vbNewLine & _
			"			<td class=""header2"" align=""center"">&nbsp;Version&nbsp;</td>" & vbNewLine & _
			"		</tr>" & vbNewLine & _
			STR_TEMP & _
			"	</table>" & vbNewLine & _
			"	</p>" & vbNewLine
	Response.Write STR_TEMP
End Sub


Sub Display_Footer()
	Response.Write  "	<p>" & vbNewLine & _
			"	<table class=""main"" cellspacing=""1"" align=""center"" style=""width:90%;"">" & vbNewLine & _
			"		<tr>" & vbNewLine & _
			"			<td class=""header"" align=""center""><a href=""http://www.onewaymule.org/onewayscripts/scripts/aspinfo/"" title=""ASPInfo"">ASPInfo v1.0</a> - Powered by <a href=""http://www.onewaymule.org/onewayscripts/"" target=""_blank"" title=""OneWayScripts - Free ASP Scripts and Snitz Forums MODs"">OneWayScripts</a></td>" & vbNewLine & _
			"	       	</tr>" & vbNewLine 			
	Response.Write  "	</table>" & vbNewLine & _
			"	</p>" & vbNewLine
End Sub


Function Display_Data(F_VARIABLE, F_VALUE)
	Dim STR_TEMP
	STR_TEMP =  "		<tr>" & vbNewLine & _
		    "			<td class=""cell"" align=""left"" nowrap>&nbsp;" & F_VARIABLE & "&nbsp;</td>" & vbNewLine & _
		    "			<td class=""cell"" align=""left"">&nbsp;" & F_VALUE & "&nbsp;</td>" & vbNewLine & _
		    "		</tr>" & vbNewLine 
	Display_Data = STR_TEMP
End Function


Response.Write  "<html>" & vbNewLine & _
		"<head>" & vbNewLine & _
		"	<meta http-equiv=""content-type"" content=""text/html; charset=utf-8"">" & vbNewLine & _
		"	<meta http-equiv=""content-language"" content=""en"">" & vbNewLine & _
		"	<meta http-equiv=""expires"" content=""0"">" & vbNewLine & _
		"	<meta name=""author"" content=""Michael Reisinger (OneWayMule)"">" & vbNewLine & _
		"	<meta name=""description"" content=""ASPInfo is a free ASP script written in VBScript which provides useful information about your Windows server, objects (Application, Request, Response, Server, Session) and installed components."">" & vbNewLine & _
		"	<meta name=""copyright"" content=""Freeware, GPL License, (c)2002-2005 Michael Reisinger (OneWayMule)"">" & vbNewLine & _
		"	<meta name=""robots"" content=""index,follow"">" & vbNewLine & _
		"	<meta name=""keywords"" content=""aspinfo,iis,asp,server,info,information,check,installed,component,components,object,objects,webserver,installed,email,test,scan"">" & vbNewLine & _
		"	<meta name=""audience"" content=""all"">" & vbNewLine & _
		"	<title>ASPInfo v1.0 - Powered by OneWayScripts</title>" & vbNewLine & _
		"	<style type=""text/css"">" & vbNewLine & _
		"		body {	font-family:verdana; font-size:13px; background-color:#FBFDFF; }" & vbNewLine & _
		"		table { padding:0px; border:0px; margin:0px; font-family:verdana; font-size:13px; color:#000000; text-align:left; }" & vbNewLine & _
		"		.main{	border: 1px solid #abb6dc; margin:1px; background-color:#ffffff; }" & vbNewLine & _
		"		.header, .header td { font-weight:bold; padding:6px; background-color:#abb6dc; color:white; }" & vbNewLine & _
		"		.header a:link,.header a:active,.header a:visited,.header a:hover { color : #FFFFFF; }" & vbNewLine & _
		"		.header2, .title  { font-weight:bold; padding:6px; background-color:#e2e7f3; color:#000000; }" & vbNewLine & _
		"		.footer  { padding:4px; font-weight:bold; background-color:#abb6dc; color:#000000; text-align:right; }" & vbNewLine & _
		"		.cell { background-color:#F1F4F7; padding:4px; }" & vbNewLine & _
		"		a:link,a:active,a:visited { color : #5c72ba; }" & vbNewLine & _
		"		a:hover { text-decoration: underline; color : #000000; }" & vbNewLine & _
		"	</style>" & vbNewLine & _
		"</head>" & vbNewLine & _
		"<body bgColor=""#ffffff"" text=""#000000"" link=""#0000ff"" aLink=""#0000ff"" vLink=""#0000ff"">" & vbNewLine & _
		"	<center>" & vbNewLine
Call Display_Header()
Call Display_Server_Information()
Call Display_Application_Object()
Call Display_Request_Object()
Call Display_Response_Object()
Call Display_Server_Object()
Call Display_Session_Object()
Call Display_All_Components()
Call Display_Footer()
Response.Write  "	</center>" & vbNewLine & _
		"</body>" & vbNewLine & _
		"</html>" & vbNewLine
%>