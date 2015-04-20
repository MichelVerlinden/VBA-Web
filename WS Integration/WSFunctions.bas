Attribute VB_Name = "WSFunctions"
'/*
'' WSFunctions
'' Module where worksheet functions should be defined
''
'' In order to define "parralel" function: Implement IAsynchWSFun
'' in a class module and use the object in AsychWSFun.asyncFun(<your object>, <your parameters>)
''
''
'' author : Michel Verlinden
'' 17/03/2014
''
'' TODO :Add generic argument validator
'' Function registration
''
'*/
Option Explicit
'Option Private Module ' comment this if not registering functions

Public Enum language
 English
 French
 German
 Spanish
 Chinese
End Enum

' Twitter
Public Function testTwitter(keyWord As String) As String
 Dim tWS As New RestWSFunction
 Dim tWeb As New TwitterQuery
 tWS.assign tWeb
 testTwitter = AsynchWSFun.asyncFun(tWS, keyWord)
End Function

' Translate
Public Function testTranslate(keyWord As String, language As String) As String
 Dim allLanguages() As String, lCodes() As String
 allLanguages = Split("Afrikaans;Akan;Albanian;Amharic;" & _
 "Arabic; Armenian;Azerbaijani; Basque; Belarusian; Bemba;Bengali; Bihari; Bork, bork, bork!;Bosnian; Breton;" & _
 "Bulgarian;Cambodian;Catalan; Cherokee;Chichewa;Chinese (Simplified);Chinese (Traditional);Corsican;Croatian;" & _
 "Czech;Danish; Dutch;Elmer Fudd; English; Esperanto;Estonian;Ewe; Faroese; Filipino;Finnish; French; Frisian;" & _
 "Ga; Galician;Georgian;German; Greek;Guarani; Gujarati;Hacker; Haitian Creole; Hausa;Hawaiian;Hebrew; Hindi;Hungarian;" & _
 "Icelandic;Igbo;Indonesian; Interlingua; Irish;Italian; Japanese;Javanese;Kannada; Kazakh; Kinyarwanda; Kirundi;" & _
 "Klingon; Kongo;Korean; Krio (Sierra Leone); Kurdish; Kurdish (Soranî);Kyrgyz; Laothian;Latin;Latvian; Lingala; Lithuanian;" & _
 "Lozi;Luganda; Luo; Macedonian; Malagasy;Malay;Malayalam;Maltese; Maori;Marathi; Mauritian Creole;Moldavian;Mongolian;Montenegrin;" & _
 "Nepali; Nigerian Pidgin; Northern Sotho; Norwegian;Norwegian (Nynorsk); Occitan; Oriya;Oromo;Pashto; Persian; Pirate; Polish; Portuguese (Brazil);" & _
 "Portuguese (Portugal);Punjabi; Quechua; Romanian;Romansh; Runyakitara; Russian; Scots Gaelic;Serbian; Serbo-Croatian; Sesotho; Setswana;" & _
 "Seychellois Creole; Shona;Sindhi; Sinhalese;Slovak; Slovenian;Somali; Spanish; Spanish (Latin American);Sundanese;Swahili; Swedish; Tajik;Tamil;" & _
 "Tatar;Telugu; Thai;Tigrinya;Tonga;Tshiluba;Tumbuka; Turkish; Turkmen; Twi; Uighur; Ukrainian;Urdu;Uzbek;Vietnamese; Welsh;Wolof;Xhosa;Yiddish; Yoruba; Zulu", ";")
 
 lCodes = Split("af; ak;sq;am;ar;hy;az;eu;be;bem;bn;bh;xx-bork;bs;br;bg;km;ca;chr;ny;zh-CN;zh-TW;co;hr;cs;da;nl;xx-elmer;en;eo;et;ee;fo;tl;fi;fr;fy;gaa;gl;ka;de;el;gn;gu;xx-hacker;ht;ha;haw;iw;hi;hu;is;ig;id;ia;ga;it;ja;jw;kn;kk;rw;rn;xx-klingon;kg;ko;kri;ku;ckb;ky;lo;la;lv;ln;lt;loz;lg;ach;mk;mg;ms;ml;mt;mi;mr;mfe;mo;mn;sr-ME;ne;pcm;nso;no;nn;oc;or;om;ps;fa;xx-pirate;pl;pt-BR;pt-PT;pa;qu;ro;rm;nyn;ru;gd;sr;sh;st;tn;crs;sn;sd;si;sk;sl;so;es;es-419;su;sw;sv;tg;ta;tt;te;th;ti;to;lua;tum;tr;tk;tw;ug;uk;ur;uz;vi;cy;wo;xh;yi;yo;zu", ";")

 Dim i As Integer
 i = 0
 While StrComp(Replace(allLanguages(i), " ", vbNullString), Replace(language, " ", vbNullString)) <> 0 _
            And i < UBound(allLanguages)
    i = i + 1
 Wend
'
 If i <= UBound(allLanguages) Then
    Dim tWS As New RestWSFunction
    Dim tWeb As New TranslateQuery
    tWeb.IRestWSQuery_Name = lCodes(i)
    tWS.assign tWeb
    testTranslate = AsynchWSFun.asyncFun(tWS, keyWord)
Else
   testTranslate = language & " is not a language"
End If
End Function


