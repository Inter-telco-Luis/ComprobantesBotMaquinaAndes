; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
SetTitleMatchMode Fast
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1


F4::
servicioPython:
pathImg := "C:\img\"
/*
pathImg := "C:\Users\Soporte\Desktop\img\"
*/
pathDescargas := "C:\Users\Soporte\Downloads\"
pathComprobantes := "C:\Users\Soporte\Desktop\Comprobantes\"
Sleep, 500
Run, http://35.175.103.14:3000/
WinWaitActive, ahk_class Chrome_WidgetWin_1
Sleep, 333
Loop
{
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200915172003.png
    CenterImgSrchCoords(pathImg "Screen_20200915172003.png", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
}
Until ErrorLevel = 0
Sleep, 300
Loop, Files, %pathComprobantes%*.*, F
{
    Sleep, 300
    If (A_LoopFileExt=="pdf")
    {
        Sleep, 300
        Loop, Parse, A_LoopFileName, %A_Space%_
        {
            If (%A_LoopField%=="rechazo" or "Rechazo")
            {
                ControlSetText, Edit1, %A_LoopFileFullPath%, ahk_class #32770
                Break
            }
        }
        Sleep, 300
    }
    /*
    MsgBox, 0, , %A_LoopFileExt%`,%A_LoopFileFullPath%
    */
}
Sleep, 300
ControlClick, Button1, ahk_class #32770,, Left, 1,  NA
Sleep, 10
Sleep, 300
Loop
{
    CoordMode, Pixel, Screen
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200929104647.png
    CenterImgSrchCoords(pathImg "Screen_20200929104647.png", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
}
Until ErrorLevel = 0
Sleep, 60000
Loop
{
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%coprobantes.PNG
    CenterImgSrchCoords(pathImg "coprobantes.PNG", FoundX, FoundY)
}
Until ErrorLevel = 0
If (ErrorLevel = 0)
{
    Loop, 3
    {
        Sleep, 300
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 10
    }
    Send, {Control Down}{c}{Control Up}
    Sleep, 3000
    /*
    Run, notepad
    Sleep, 2000
    Send, 
    Send, {Control Down}{v}{Control Up}
    Sleep, 3000
    Send, 
    Send, {Control Down}{g}{Control Up}
    Sleep, 2000
    ControlSetText, Edit1, %pathComprobantes%result.json, ahk_class #32770
    Sleep, 2000
    ControlClick, Button2, ahk_class #32770,, Left, 1,  NA
    Sleep, 10
    Sleep, 2000
    */
    Send, {Alt Down}{F4}{Alt Up}
    Sleep, 1000
}
Goto, Macro7
Return

Macro7:
/*
pathImg := "C:\Users\Soporte\Desktop\img\"
*/
pathImg := "C:\img\"
pathDescargas := "C:\Users\Soporte\Downloads\"
pathComprobantes := "C:\Users\Soporte\Desktop\Comprobantes\"
/*
result := ""
FileRead, result, %pathComprobantes%result.json
*/
banImg := 0
abonados := 1
cuenta := ""
beneficiario := ""
cc := ""
valor := ""
img := 0
referencia := ""
page := ""
resultArray := []
Loop, Parse, Clipboard, `,., "{}
{
    Loop, Parse, A_LoopField, :`,, "
    {
        /*
        MsgBox, 0, , %A_LoopField%
        */
        resultArray.Push(A_LoopField)
    }
}
/*
Sleep, 1000
Send, {Alt Down}{F4}{Alt Up}
*/
Sleep, 1000
loginDynamics(pathImg)
Sleep, 500
/*
indexItem := 0
*/
For key, value in resultArray
{
    MsgBox, 0, , %value%`,%key%
    If (value == "No_Abonados")
    {
        abonados := 0
    }
    Else If (value == "numero de cuenta")
    {
        cuenta := resultArray[ "A_Index" +1]
    }
    Else If (value == "nombre de beneficiario")
    {
        beneficiario := resultArray[ "A_Index" +1]
    }
    Else If (value == "documento")
    {
        cc := resultArray[ "A_Index" +1]
    }
    Else If (value == "valor")
    {
        valor := resultArray[ "A_Index" +1]
        valor .= "."
        valor .= resultArray[ "A_Index" +2]
    }
    Else If (value == "referencia")
    {
        referencia := resultArray[ "A_Index" +1]
    }
    Else If (value == "img")
    {
        banImg := 1
    }
    Else If (banImg == "1")
    {
        img := value
        banImg := 0
    }
    Else If (value == "page")
    {
        page := resultArray[ "A_Index" +1]
        Sleep, 1000
        /*
        MsgBox, 0, , %pathImg%
        Sleep, 1000
        */
        auxAbonado(cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,abonados,pathComprobantes)
        Sleep, 1000
        /*
        If (img != 9223372036854775807)
        {
        }
        */
    }
}
/*
MsgBox, 0, , 
(LTrim
% resultArray[%indexItem%+1],% resultArray[%indexItem%+2]

)
MsgBox, 0, , %cc%`,%valor%`,%referencia%
img .= resultArray[ "A_Index" +2]
*/
Return

b64Jpg(img, pathImg, pathDescargas)
{
    Sleep, 1000
    /*
    MsgBox, 0, , %pathDescargas%
    */
    FileDelete, %pathDescargas%test-download.png
    Sleep, 1000
    Run, http://35.175.103.14:5000/
    Sleep, 500
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201022084558.png
        CenterImgSrchCoords(pathImg "Screen_20201022084558.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 500
    Clipboard := "images/" img ".png"
    Sleep, 1000
    Send, {Control Down}{v}{Control Up}
    Sleep, 1000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201022084614.png
        CenterImgSrchCoords(pathImg "Screen_20201022084614.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 1000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201022084714.png
        CenterImgSrchCoords(pathImg "Screen_20201022084714.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 3500
    Send, {Alt Down}{F4}{Alt Up}
    Sleep, 500
}

loginDynamics(ByRef pathImg)
{
    If (!IsObject(ie))
        ie := ComObjCreate("InternetExplorer.Application")
    ie.Visible := true
    ie.Navigate("https://gco.crm2.dynamics.com/")
    IELoad(ie)
    Sleep, 2000
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201021164928.png
    CenterImgSrchCoords(pathImg "Screen_20201021164928.png", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
    Sleep, 2000
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\Soporte\Desktop\img\iniciar.PNG
    If (ErrorLevel)
    {
        Send, {F5}
        Sleep, 2000
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200922121429.png
            CenterImgSrchCoords(pathImg "Screen_20200922121429.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
    }
    Sleep, 3000
    Send, {a}
    Sleep, 500
    Send, {Backspace}
    Sleep, 500
    userDynamics := "agentecrm2.ext@gco.com.co"
    passwordDynamics := "Ag3nt3_2018*"
    ie.document.getElementsByTagName("INPUT")[0].InnerText := userDynamics
    Sleep, 500
    Send, {Enter}
    /*
    ie.document.getElementsByTagName("INPUT")[3].Click("")
    */
    Sleep, 1000
    Send, {b}{Backspace}
    ie.document.getElementsByTagName("INPUT")[9].InnerText := passwordDynamics
    ie.document.getElementsByTagName("INPUT")[10].Click("")
    Sleep, 2000
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200922110820.png
    If (ErrorLevel = 0)
    {
        /*
        Sleep, 1000
        */
        ie.document.getElementsByTagName("INPUT")[6].Click("")
        /*
        Sleep, 1000
        */
        ie.document.getElementsByTagName("INPUT")[7].Click("")
    }
    Sleep, 5000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200916211549.png
        CenterImgSrchCoords(pathImg "Screen_20200916211549.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 2000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png
        CenterImgSrchCoords(pathImg "Screen_20200918143624.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 2000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201001114340.png
        CenterImgSrchCoords(pathImg "Screen_20201001114340.png", FoundX, FoundY)
    }
    Until ErrorLevel = 0
    If (ErrorLevel = 0)
    {
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 10
    }
    /*
    MsgBox, 0, , %codhtml%
    */
    Sleep, 4000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200921213437.png
        CenterImgSrchCoords(pathImg "Screen_20200921213437.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 2000
    Send, {End}
    Sleep, 2000
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200921213703.png
    CenterImgSrchCoords(pathImg "Screen_20200921213703.png", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
    If (ErrorLevel)
    {
        Sleep, 2000
        Send, {End}
        Sleep, 2000
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200921213703.png
        CenterImgSrchCoords(pathImg "Screen_20200921213703.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Sleep, 1000
}

auxSarch(cuenta, beneficiario, cc, valor, referencia, img, page, pathImg, pathDescargas, pathComprobantes)
{
    Sleep, 500
    MsgBox, 0, , estoy en AuxSearch y la cedula es: %cc%
    comprobante := searchDynamics(cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,pathComprobantes)
    Sleep, 3000
    /*
    MsgBox, 0, , %comprobante%
    */
    If (comprobante == 1)
    {
        FileAppend, `n  nombre:%beneficiario%.Documento:%cc%`,Valor:%valor%`,pagina:%page%hora:%A_Tab%%A_Hour%:%A_Min%:%A_Sec%%A_Tab%, %pathComprobantes%enviados.txt
        Sleep, 300
    }
    Else
    {
        auxComprobante := searchDynamics(cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,pathComprobantes)
        Sleep, 3000
        If (auxComprobante == 0)
        {
            FileAppend, `n  nombre:%beneficiario%.Documento:%cc%`,Valor:%valor%`,pagina:%page%hora:%A_Tab%%A_Hour%:%A_Min%:%A_Sec%%A_Tab%, %pathComprobantes%inconvenientes.txt
        }
        Else
        {
            FileAppend, `n  nombre:%beneficiario%.Documento:%cc%`,Valor:%valor%`,pagina:%page%hora:%A_Tab%%A_Hour%:%A_Min%:%A_Sec%%A_Tab%, %pathComprobantes%enviados.txt
        }
    }
    return comprobante
}

auxAbonado(cuenta, beneficiario, cc, valor, referencia, img, page, pathImg, pathDescargas, abonados, pathComprobantes)
{
    MsgBox, 0, , %cc% en auxAbonado
    If (abonados == 1)
    {
        /*
        MsgBox, 0, , Lugar de llamada a la funcion de busqueda de cedula y envio de correo
        */
        Sleep, 1000
        comprobante := auxSarch(cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,pathComprobantes)
        Sleep, 1000
    }
    Else
    {
        FileAppend, `n  nombre:%beneficiario%.Documento:%cc%`,Valor:%valor%`,pagina:%page%hora:%A_Tab%%A_Hour%:%A_Min%:%A_Sec%%A_Tab%, %pathComprobantes%inconvenientes.txt
    }
    /*
    imagen.Push(% resultArray[%A_Index%+1])
    imagen.Push(resultArray[A_Index+1])
    banImg := 1
}
Else If (banImg = 1)
{
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201008134856.png
    CenterImgSrchCoords(pathImg "Screen_20201008134856.png", FoundX, FoundY)
    Click, %FoundX%, %FoundY% Left, 1
    Sleep, 10
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201008135352.png
    CenterImgSrchCoords(pathImg "Screen_20201008135352.png", FoundX, FoundY)
    Sleep, 500
    If (ErrorLevel = 0)
    {
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 10
    }
    Sleep, 1000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png
        CenterImgSrchCoords(pathImg "Screen_20200918143624.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 1500
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201007154153.png
    CenterImgSrchCoords(pathImg "Screen_20201007154153.png", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
    Sleep, 1000
    If (ErrorLevel = 0)
    {
        Send, {Enter}
        Sleep, 500
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201007151603.png
            CenterImgSrchCoords(pathImg "Screen_20201007151603.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
        Sleep, 500
        Clipboard := "Resuelto."
        Sleep, 300
        Send, {Control Down}{v}{Control Up}
        Sleep, 500
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20200924082550.png
            CenterImgSrchCoords(pathImg "Screen_20200924082550.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
        Sleep, 2000
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201007151912.png
            CenterImgSrchCoords(pathImg "Screen_20201007151912.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
        Sleep, 3000
    }
    Else
    {
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20200924082550.png
            CenterImgSrchCoords(pathImg "Screen_20200924082550.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
    }
    Sleep, 2000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200929114218.png
        CenterImgSrchCoords(pathImg "Screen_20200929114218.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 2000
    IfWinActive, Mensaje de página web
    {
        ControlClick, Button1, Mensaje de página web,, Left, 1,  NA
        Sleep, 10
        Sleep, 1000
    }
    Click, %FoundX%, %FoundY% Left, 1
    Sleep, 10
    Sleep, 1000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png
        CenterImgSrchCoords(pathImg "Screen_20200918143624.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 2000
}
}
Else
{
status := 0
Sleep, 3000
Loop
{
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20200923182420.png
    CenterImgSrchCoords(pathImg "Screen_20200923182420.png", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
}
Until ErrorLevel = 0
Sleep, 10000
Loop
{
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200924143844.png
    CenterImgSrchCoords(pathImg "Screen_20200924143844.png", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
}
Until ErrorLevel = 0
Sleep, 1000
Loop
{
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png
    CenterImgSrchCoords(pathImg "Screen_20200918143624.png", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
}
Until ErrorLevel = 0
Sleep, 1000
}
return status
*/
}

searchDynamics(cuenta, beneficiario, cc, valor, referencia, img, page, pathImg, pathDescargas, pathComprobantes)
{
    global
    MsgBox, 0, ,  estoy en searchDinamycs y esta es la cedula %cc%
    status := 1
    Sleep, 6000
    CoordMode, Pixel, Screen
    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201022080131.png
    If (ErrorLevel)
    {
        Sleep, 2000
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022080228.png
            CenterImgSrchCoords(pathImg "Screen_20201022080228.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
        Send, {End}
        auxComprobante := searchDynamics(cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,pathComprobantes)
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022095157.png
            CenterImgSrchCoords(pathImg "Screen_20201022095157.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
    }
    Sleep, 3000
    /*
    MsgBox, 0, , Antes de Maximizar
    WinMaximize, Casos Vista de devoluciones de dinero - Microsoft Dynamics 365 - Internet Explorer
    Sleep, 333
    */
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022151658.png
        CenterImgSrchCoords(pathImg "Screen_20201022151658.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 5000
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%filterWhite.PNG
    CenterImgSrchCoords(pathImg "filterWhite.PNG", FoundX, FoundY)
    If (ErrorLevel = 0)
    {
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 10
        Sleep, 2000
        Click, 0, 0, 0
        Sleep, 10
    }
    /*
    MsgBox, 0, , buscando filtro negro
    */
    Sleep, 1000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%blackFilter.PNG
        CenterImgSrchCoords(pathImg "blackFilter.PNG", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    /*
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022080608.png
        CenterImgSrchCoords(pathImg "Screen_20201022080608.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    */
    Sleep, 2000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022080712.png
        CenterImgSrchCoords(pathImg "Screen_20201022080712.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 500
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022080759.png
        CenterImgSrchCoords(pathImg "Screen_20201022080759.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 500
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%aceptarDD.PNG
        CenterImgSrchCoords(pathImg "aceptarDD.PNG", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 2000
    Loop, 3
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022080849.png
        CenterImgSrchCoords(pathImg "Screen_20201022080849.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 3000
    Send, {Tab}
    Sleep, 1000
    Send, {Enter}
    Sleep, 5000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201022080916.png
        CenterImgSrchCoords(pathImg "Screen_20201022080916.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 1000
    Loop, 5
    {
        Send, {Right}
        Sleep, 1000
    }
    startX := 0
    Loop, 7
    {
        Loop, 3
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, %startX%, 0, 1920, 1080, %pathImg%Screen_20201022080940.png
            CenterImgSrchCoords(pathImg "Screen_20201022080940.png", FoundX, FoundY)
        }
        Until ErrorLevel = 0
        startX := FoundX
    }
    Click, %FoundX%, %FoundY% Left, 1
    Sleep, 10
    Sleep, 300
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022081032.png
        CenterImgSrchCoords(pathImg "Screen_20201022081032.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 2000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022081251.png
        CenterImgSrchCoords(pathImg "Screen_20201022081251.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 2000
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%selectcc.PNG
    Sleep, 1000
    If (ErrorLevel = 0)
    {
        bccfiltro := 1
        Sleep, 1000
    }
    Send, {Down}
    Sleep, 500
    Send, {Enter}
    Sleep, 500
    Send, {Tab}
    Sleep, 1000
    Clipboard := cc
    Sleep, 1000
    /*
    MsgBox, 0, , %cc%`, Esta es la cedula
    */
    Send, {Control Down}{v}{Control Up}
    Sleep, 1000
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%aceptarCC.PNG
        CenterImgSrchCoords(pathImg "aceptarCC.PNG", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    Sleep, 4000
    /*
    MsgBox, 0, , Buscanco "Bu"
    */
    Loop
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201103103155.png
        CenterImgSrchCoords(pathImg "Screen_20201103103155.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Until ErrorLevel = 0
    If (ErrorLevel)
    {
        Sleep, 1000
        /*
        MsgBox, 0, , Buscando Esquina
        */
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022082733.png
        CenterImgSrchCoords(pathImg "Screen_20201022082733.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
    }
    Sleep, 2000
    Send, {Control Down}{f}{Control Up}
    Sleep, 2000
    Clipboard := valor
    Sleep, 1000
    Send, {Control Down}{v}{Control Up}
    Sleep, 1000
    Send, {Enter}
    Sleep, 2000
    /*
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022083430.png
    CenterImgSrchCoords(pathImg "Screen_20201022083430.png", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%encuentroValor.PNG
    CenterImgSrchCoords(pathImg "encuentroValor.PNG", FoundX, FoundY)
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
    */
    CoordMode, Pixel, Window
    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%findValue.PNG
    If ErrorLevel = 0
    	Click, %FoundX%, %FoundY% Left, 1
    Sleep, 200
    /*
    MsgBox, 0, , %bccfiltro%`,%ErrorLevel%
    */
    If (bccfiltro==1 && ErrorLevel==0)
    {
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 10
        Sleep, 10000
        Send, {Control Down}{f}{Control Up}
        Sleep, 2000
        CoordMode, Pixel, Window
        PixelSearch, FoundX, FoundY, 0, 0, 1280, 1024, 0x0078D7, 0, Fast RGB
        If (ErrorLevel)
        {
            Send, {Control Down}{f}{Control Up}
            Sleep, 2000
        }
        Clipboard := "Agregar Llama"
        Sleep, 2000
        Send, {Control Down}{v}{Control Up}
        Sleep, 2000
        startY := 0
        Loop, 2
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, %startY%, 1920, 1080, %pathImg%Screen_20201022083533.png
            CenterImgSrchCoords(pathImg "Screen_20201022083533.png", FoundX, FoundY)
            startY := FoundY
        }
        Sleep, 500
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 10
        /*
        MsgBox, 0, , %ErrorLevel%
        */
        If (ErrorLevel)
        {
            status := 0
            Sleep, 3000
            If (!IsObject(ie))
                ie := ComObjCreate("InternetExplorer.Application")
            ie.Visible := true
            ie.Navigate("https://gco.crm2.dynamics.com/main.aspx")
            IELoad(ie)
            /*
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201022083610.png
                CenterImgSrchCoords(pathImg "Screen_20201022083610.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            */
            Sleep, 10000
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022094907.png
                CenterImgSrchCoords(pathImg "Screen_20201022094907.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 1000
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022083735.png
                CenterImgSrchCoords(pathImg "Screen_20201022083735.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
        }
        If (%status%==1)
        {
            Sleep, 1000
            Send, {Enter}
            Sleep, 1000
            Clipboard := ""
            Sleep, 2000
            /*
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022083826.png
                CenterImgSrchCoords(pathImg "Screen_20201022083826.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 2000
            Send, {Control Down}{c 2}{Control Up}
            Sleep, 2000
            Send, {Control Down}{c}{Control Up}
            Sleep, 2000
            IfWinActive, ahk_class #32770
            {
                ControlClick, Button2, ahk_class #32770,, Left, 1,  NA
                Sleep, 10
                Sleep, 2000
            }
            If (Clipboard = "")
            {
                */
                Clipboard := "Radicado Devolucion de Dinero"
                /*
                MsgBox, 0, , %Clipboard%
            }
            */
            Sleep, 500
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022083846.png
                CenterImgSrchCoords(pathImg "Screen_20201022083846.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 2000
            Send, {Control Down}{v}{Control Up}
            Sleep, 3000
            Loop, 7
            {
                Send, {Tab}
                Sleep, 100
            }
            Clipboard := 
            (LTrim
            "Hola, 
            
            Queremos contarte que la devolución del dinero ya fue efectiva, te enviamos el comprobante de transferencia.
            
            Cualquier inquietud adicional, con gusto será atendida.
            
            SERVICIO AL CLIENTE."
            )
            Send, {Control Down}{v}{Control Up}
            Sleep, 500
            /*
            MsgBox, 0, , %pathImg%`,%pathDescargas%
            */
            b64Jpg(img,pathImg,pathDescargas)
            Sleep, 1000
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022083904.png
                CenterImgSrchCoords(pathImg "Screen_20201022083904.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 2000
            Send, {Alt Down}{Tab}{Alt Up}
            Sleep, 500
            Send, {Alt Down}{Tab}{Alt Up}
            Sleep, 500
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022085611.png
                CenterImgSrchCoords(pathImg "Screen_20201022085611.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 2000
            ControlSetText, Edit1, %pathDescargas%test-download.png, ahk_class #32770
            Sleep, 4000
            ControlClick, Button1, ahk_class #32770,, Left, 1,  NA
            Sleep, 10
            Sleep, 1000
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022113123.png
                CenterImgSrchCoords(pathImg "Screen_20201022113123.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 3000
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022113437.png
                CenterImgSrchCoords(pathImg "Screen_20201022113437.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 4000
            /*
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201022084057.png
                CenterImgSrchCoords(pathImg "Screen_20201022084057.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 6000
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20200924033234.png
            CenterImgSrchCoords(pathImg "Screen_20200924033234.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
            Sleep, 2000
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201008135352.png
            CenterImgSrchCoords(pathImg "Screen_20201008135352.png", FoundX, FoundY)
            Sleep, 500
            If (ErrorLevel = 0)
            {
                Click, %FoundX%, %FoundY% Left, 1
                Sleep, 10
            }
            Sleep, 1000
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png
                CenterImgSrchCoords(pathImg "Screen_20200918143624.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 1500
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201007154153.png
            CenterImgSrchCoords(pathImg "Screen_20201007154153.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
            Sleep, 1000
            If (ErrorLevel = 0)
            {
                Send, {Enter}
                Sleep, 500
                Loop
                {
                    CoordMode, Pixel, Window
                    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201007151603.png
                    CenterImgSrchCoords(pathImg "Screen_20201007151603.png", FoundX, FoundY)
                    If ErrorLevel = 0
                    	Click, %FoundX%, %FoundY% Left, 1
                }
                Until ErrorLevel = 0
                Sleep, 500
                Clipboard := "Resuelto."
                Sleep, 300
                Send, {Control Down}{v}{Control Up}
                Sleep, 500
                Loop
                {
                    CoordMode, Pixel, Window
                    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20200924082550.png
                    CenterImgSrchCoords(pathImg "Screen_20200924082550.png", FoundX, FoundY)
                    If ErrorLevel = 0
                    	Click, %FoundX%, %FoundY% Left, 1
                }
                Until ErrorLevel = 0
                Sleep, 2000
                Loop
                {
                    CoordMode, Pixel, Window
                    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20201007151912.png
                    CenterImgSrchCoords(pathImg "Screen_20201007151912.png", FoundX, FoundY)
                    If ErrorLevel = 0
                    	Click, %FoundX%, %FoundY% Left, 1
                }
                Until ErrorLevel = 0
                Sleep, 3000
            }
            Else
            {
                Loop
                {
                    CoordMode, Pixel, Window
                    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20200924082550.png
                    CenterImgSrchCoords(pathImg "Screen_20200924082550.png", FoundX, FoundY)
                    If ErrorLevel = 0
                    	Click, %FoundX%, %FoundY% Left, 1
                }
                Until ErrorLevel = 0
            }
            Sleep, 2000
            MsgBox, 0, , inicia Buscque de imagen X
            */
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%xmessage.PNG
                CenterImgSrchCoords(pathImg "xmessage.PNG", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 2000
            /*
            IfWinActive, Mensaje de página web
            {
                */
                Loop
                {
                    CoordMode, Pixel, Window
                    ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%aceptarMesX.PNG
                    CenterImgSrchCoords(pathImg "aceptarMesX.PNG", FoundX, FoundY)
                    If ErrorLevel = 0
                    	Click, %FoundX%, %FoundY% Left, 1
                }
                Until ErrorLevel = 0
                /*
                ControlClick, Button1, Mensaje de página web,, Left, 1,  NA
                Sleep, 10
                */
                Sleep, 1000
                /*
            }
            */
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%xmessage.PNG
                CenterImgSrchCoords(pathImg "xmessage.PNG", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            /*
            Click, %FoundX%, %FoundY% Left, 1
            Sleep, 10
            */
            Sleep, 1000
            /*
            MsgBox, 0, , a continuacion se minimiza pestaña
            */
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png
                CenterImgSrchCoords(pathImg "Screen_20200918143624.png", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 2000
        }
    }
    Else
    {
        status := 0
        Sleep, 3000
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, %pathImg%Screen_20200923182420.png
            CenterImgSrchCoords(pathImg "Screen_20200923182420.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
        Sleep, 10000
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200924143844.png
            CenterImgSrchCoords(pathImg "Screen_20200924143844.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
        Sleep, 1000
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png
            CenterImgSrchCoords(pathImg "Screen_20200918143624.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
        Sleep, 1000
    }
    return status
}


IELoad(Pwb)
{
	While (!Pwb.busy)
		Sleep, 100
	While (Pwb.busy)
		Sleep, 100
	While (!Pwb.document.Readystate = "Complete")
		Sleep, 100
}

CenterImgSrchCoords(File, ByRef CoordX, ByRef CoordY)
{
	static LoadedPic
	LastEL := ErrorLevel

	Gui, Pict:Add, Pic, vLoadedPic, % RegExReplace(File, "^(\*\w+\s)+")
	GuiControlGet, LoadedPic, Pict:Pos
	Gui, Pict:Destroy
	CoordX += LoadedPicW // 2
	CoordY += LoadedPicH // 2
	ErrorLevel := LastEL
}

/*
PMC File Version 5.3.4
---[Do not edit anything in this section]---

[PMC Code v5.3.4]|F4||1|Window,2,Fast,0,1,Input,-1,-1,1|1|servicioPython
Context=None||None|
Groups=Start:1
1|[Assign Variable]|pathImg := C:\img\|1|0|Variable|||||1|
02|[Assign Variable]|pathImg := C:\Users\Soporte\Desktop\img\|1|0|Variable|||||2|
3|[Assign Variable]|pathDescargas := C:\Users\Soporte\Downloads\|1|0|Variable|||||4|
4|[Assign Variable]|pathComprobantes := C:\Users\Soporte\Desktop\Comprobantes\|1|0|Variable|||||6|
5|[Pause]||1|500|Sleep|||||7|
6|Run|http://35.175.103.14:3000/|1|0|Run|||||8|
7|WinWaitActive||1|333|WinWaitActive||ahk_class Chrome_WidgetWin_1|||9|
8|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200915172003.png|1|0|ImageSearch|UntilFound|Window|||11|
9|[Pause]||1|300|Sleep|||||20|
10|[LoopStart]|%pathComprobantes%*.*`, F|1|0|LoopFilePattern|||||21|
11|[Pause]||1|300|Sleep|||||23|
12|Evaluate Expression|A_LoopFileExt=="pdf"|1|0|If_Statement|0||||24|
13|[Pause]||1|300|Sleep|||||26|
14|[LoopStart]|A_LoopFileName`, %A_Space%_`, |1|0|LoopParse|||||27|
15|Evaluate Expression|%A_LoopField%=="rechazo" or "Rechazo"|1|0|If_Statement|||||29|
16|[Control]|%A_LoopFileFullPath%|1|0|ControlSetText|Edit1|ahk_class #32770|||31|
17|Break||1|0|Break|||||32|
18|[End If]|EndIf|1|0|If_Statement|||||33|
19|[LoopEnd]|LoopEnd|1|0|Loop|||||34|
20|[Pause]||1|300|Sleep|||||35|
21|[End If]|EndIf|1|0|If_Statement|||||36|
022|[MsgBox]|%A_LoopFileExt%`,%A_LoopFileFullPath%|1|0|MsgBox|0||||37|
23|[LoopEnd]|LoopEnd|1|0|Loop|||||39|
24|[Pause]||1|300|Sleep|||||41|
25|Left Click|Left, 1,  NA|1|10|ControlClick|Button1|ahk_class #32770|||42|
26|[Pause]||1|300|Sleep|||||44|
27|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200929104647.png|1|0|ImageSearch|UntilFound|Screen|||45|
28|[Pause]||1|60000|Sleep|||||54|
29|Continue, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%coprobantes.PNG|1|0|ImageSearch|UntilFound|Window|||55|
30|If Image/Pixel Found||1|0|If_Statement|||||62|
31|[LoopStart]|LoopStart|3|0|Loop|||||64|
32|[Pause]||1|300|Sleep|||||66|
33|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||67|
34|[LoopEnd]|LoopEnd|1|0|Loop|||||69|
35|Control + c|{Control Down}{c}{Control Up}|1|0|Send|||||70|
36|[Pause]||1|3000|Sleep|||||71|
037|Run|notepad|1|0|Run|||||72|
038|[Pause]||1|2000|Sleep|||||74|
039|Control + v|{Control Down}{v}{Control Up}|1|0|Send|||||75|
040|[Pause]||1|3000|Sleep|||||77|
041|Control + g|{Control Down}{g}{Control Up}|1|0|Send|||||78|
042|[Pause]||1|2000|Sleep|||||80|
043|[Control]|%pathComprobantes%result.json|1|0|ControlSetText|Edit1|ahk_class #32770|||81|
044|[Pause]||1|2000|Sleep|||||82|
045|Left Click|Left, 1,  NA|1|10|ControlClick|Button2|ahk_class #32770|||83|
046|[Pause]||1|2000|Sleep|||||85|
47|Alt + F4|{Alt Down}{F4}{Alt Up}|1|0|Send|||||86|
48|[Pause]||1|1000|Sleep|||||88|
49|[End If]|EndIf|1|0|If_Statement|||||89|
50|[Goto]|Macro7|1|0|Goto|||||90|

[PMC Code v5.3.4]|||1|Window,2,Fast,0,1,Input,-1,-1,1|1|Macro7
Context=None||None|
Groups=Start:1
01|[Assign Variable]|pathImg := C:\Users\Soporte\Desktop\img\|1|0|Variable|||||1|
2|[Assign Variable]|pathImg := C:\img\|1|0|Variable|||||3|
3|[Assign Variable]|pathDescargas := C:\Users\Soporte\Downloads\|1|0|Variable|||||5|
4|[Assign Variable]|pathComprobantes := C:\Users\Soporte\Desktop\Comprobantes\|1|0|Variable|||||6|
05|[Assign Variable]|result := ""|1|0|Variable|||||7|
06|FileRead|result, %pathComprobantes%result.json|1|0|FileRead|||||9|
7|[Assign Variable]|banImg := 0|1|0|Variable|||||10|
8|[Assign Variable]|abonados := 1|1|0|Variable|||||12|
9|[Assign Variable]|cuenta := ""|1|0|Variable|||||13|
10|[Assign Variable]|beneficiario := ""|1|0|Variable|||||14|
11|[Assign Variable]|cc := ""|1|0|Variable|||||15|
12|[Assign Variable]|valor := ""|1|0|Variable|||||16|
13|[Assign Variable]|img := 0|1|0|Variable|||||17|
14|[Assign Variable]|referencia := ""|1|0|Variable|||||18|
15|[Assign Variable]|page := ""|1|0|Variable|||||19|
16|[Assign Variable]|resultArray := []|1|0|Variable|Expression||||20|
17|[LoopStart]|Clipboard`, ``,.`, "{}|1|0|LoopParse|||||21|
18|[LoopStart]|A_LoopField`, :``,`, "|1|0|LoopParse|||||23|
019|[MsgBox]|%A_LoopField%|1|0|MsgBox|0||||25|
20|Push|_null := A_LoopField|1|0|Method|resultArray||||27|
21|[LoopEnd]|LoopEnd|1|0|Loop|||||29|
22|[LoopEnd]|LoopEnd|1|0|Loop|||||30|
023|[Pause]||1|1000|Sleep|||||31|
024|Alt + F4|{Alt Down}{F4}{Alt Up}|1|0|Send|||||33|
25|[Pause]||1|1000|Sleep|||||34|
26|loginDynamics|_null := pathImg|1|0|Function|||||36|
27|[Pause]||1|500|Sleep|||||37|
028|[Assign Variable]|indexItem := 0|1|0|Variable|||||38|
29|[LoopStart]|resultArray`, key`, value|1|0|For|||||40|
30|[MsgBox]|%value%`,%key%|1|0|MsgBox|0||||43|
31|Compare Variables|value == "No_Abonados"|1|0|If_Statement|||||44|
32|[Assign Variable]|abonados := 0|1|0|Variable|||||46|
33|[ElseIf] Compare Variables|value == "numero de cuenta"|1|0|If_Statement|||||47|
34|[Assign Variable]|cuenta := % resultArray[%A_Index%+1]|1|0|Variable|||||50|
35|[ElseIf] Compare Variables|value == "nombre de beneficiario"|1|0|If_Statement|||||51|
36|[Assign Variable]|beneficiario := % resultArray[%A_Index%+1]|1|0|Variable|||||54|
37|[ElseIf] Compare Variables|value == "documento"|1|0|If_Statement|||||55|
38|[Assign Variable]|cc := % resultArray[%A_Index%+1]|1|0|Variable|||||58|
39|[ElseIf] Compare Variables|value == "valor"|1|0|If_Statement|||||59|
40|[Assign Variable]|valor := % resultArray[%A_Index%+1]|1|0|Variable|||||62|
41|[Concatenate Variable]|valor .= .|1|0|Variable|||||63|
42|[Concatenate Variable]|valor .= % resultArray[%A_Index%+2]|1|0|Variable|||||64|
43|[ElseIf] Compare Variables|value == "referencia"|1|0|If_Statement|||||65|
44|[Assign Variable]|referencia := % resultArray[%A_Index%+1]|1|0|Variable|||||68|
45|[ElseIf] Compare Variables|value == "img"|1|0|If_Statement|||||69|
46|[Assign Variable]|banImg := 1|1|0|Variable|||||72|
47|[ElseIf] Compare Variables|banImg == "1"|1|0|If_Statement|||||73|
48|[Assign Variable]|img := %value%|1|0|Variable|||||76|
49|[Assign Variable]|banImg := 0|1|0|Variable|||||77|
50|[ElseIf] Compare Variables|value == "page"|1|0|If_Statement|||||78|
51|[Assign Variable]|page := % resultArray[%A_Index%+1]|1|0|Variable|||||81|
52|[Pause]||1|1000|Sleep|||||82|
053|[MsgBox]|%pathImg%|1|0|MsgBox|0||||83|
054|[Pause]||1|1000|Sleep|||||85|
55|auxAbonado|_null := cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,abonados,pathComprobantes|1|0|Function|||||86|
56|[Pause]||1|1000|Sleep|||||88|
057|Evaluate Expression|img != 9223372036854775807|1|0|If_Statement|||||89|
058|[End If]|EndIf|1|0|If_Statement|||||92|
59|[End If]|EndIf|1|0|If_Statement|||||93|
60|[LoopEnd]|LoopEnd|1|0|Loop|||||95|
061|[MsgBox]|% resultArray[%indexItem%+1]`,% resultArray[%indexItem%+2]`n|1|0|MsgBox|0||||96|
062|[MsgBox]|%cc%`,%valor%`,%referencia%|1|0|MsgBox|0||||102|
063|[Concatenate Variable]|img .= % resultArray[%A_Index%+2]|1|0|Variable|||||103|

[PMC Code v5.3.4]|||1|Window,2,Fast,0,1,Input,-1,-1,1|1|b64Jpg()
Context=None||None|
Groups=Start:1
1|[FuncParameter]|img|1|0|FuncParameter|||||1|
2|[FuncParameter]|pathImg|1|0|FuncParameter|||||1|
3|[FuncParameter]|pathDescargas|1|0|FuncParameter|||||1|
4|[FunctionStart]|b64Jpg|1|0|UserFunction|Local| / |||1|
5|[Pause]||1|1000|Sleep|||||3|
06|[MsgBox]|%pathDescargas%|1|0|MsgBox|0||||4|
7|FileDelete|%pathDescargas%test-download.png|1|0|FileDelete|||||6|
8|[Pause]||1|1000|Sleep|||||8|
9|Run|http://35.175.103.14:5000/|1|0|Run|||||9|
10|[Pause]||1|500|Sleep|||||10|
11|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201022084558.png|1|0|ImageSearch|UntilFound|Window|||11|
12|[Pause]||1|500|Sleep|||||20|
13|[Assign Variable]|Clipboard := images/%img%.png|1|0|Variable|||||21|
14|[Pause]||1|1000|Sleep|||||22|
15|Control + v|{Control Down}{v}{Control Up}|1|0|Send|||||23|
16|[Pause]||1|1000|Sleep|||||24|
17|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201022084614.png|1|0|ImageSearch|UntilFound|Window|||25|
18|[Pause]||1|1000|Sleep|||||34|
19|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201022084714.png|1|0|ImageSearch|UntilFound|Window|||35|
20|[Pause]||1|3500|Sleep|||||44|
21|Alt + F4|{Alt Down}{F4}{Alt Up}|1|0|Send|||||45|
22|[Pause]||1|500|Sleep|||||46|

[PMC Code v5.3.4]|||1|Window,2,Fast,0,1,Input,-1,-1,1|1|loginDynamics()
Context=None||None|
Groups=Start:1
1|[FuncParameter]|pathImg|1|0|FuncParameter|ByRef||||1|
2|[FunctionStart]|loginDynamics|1|0|UserFunction|Local| / |||1|
3|Method:Navigate:|https://gco.crm2.dynamics.com/|1|0|IECOM_Set|:|LoadWait|||3|
4|[Pause]||1|2000|Sleep|||||8|
5|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201021164928.png|1|0|ImageSearch||Window|||9|
6|[Pause]||1|2000|Sleep|||||14|
7|Continue, Continue, FoundX, FoundY, 0|0, 0, 1920, 1080, C:\Users\Soporte\Desktop\img\iniciar.PNG|1|0|ImageSearch||Window|||15|
8|If Image/Pixel Not Found||1|0|If_Statement|||||17|
9|F5|{F5}|1|0|Send|||||19|
10|[Pause]||1|2000|Sleep|||||20|
11|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200922121429.png|1|0|ImageSearch|UntilFound|Window|||21|
12|[End If]|EndIf|1|0|If_Statement|||||30|
13|[Pause]||1|3000|Sleep|||||31|
14|a|{a}|1|0|Send|||||32|
15|[Pause]||1|500|Sleep|||||33|
16|Backspace|{Backspace}|1|0|Send|||||34|
17|[Pause]||1|500|Sleep|||||35|
18|[Assign Variable]|userDynamics := agentecrm2.ext@gco.com.co|1|0|Variable|||||36|
19|[Assign Variable]|passwordDynamics := Ag3nt3_2018*|1|0|Variable|||||37|
20|Property:InnerText:TagName|%userDynamics%|1|0|IECOM_Set|INPUT:0||||38|
21|[Pause]||1|500|Sleep|||||39|
22|Enter|{Enter}|1|0|Send|||||40|
023|Method:Click:TagName||1|0|IECOM_Set|INPUT:3||||41|
24|[Pause]||1|1000|Sleep|||||43|
25|b|{b}|1|0|Send|||||45|
26|Backspace|{Backspace}|1|0|Send|||||46|
27|Property:InnerText:TagName|%passwordDynamics%|1|0|IECOM_Set|INPUT:9||||46|
28|Method:Click:TagName||1|0|IECOM_Set|INPUT:10||||47|
29|[Pause]||1|2000|Sleep|||||48|
30|Continue, Continue, FoundX, FoundY, 0|0, 0, 1920, 1080, %pathImg%Screen_20200922110820.png|1|0|ImageSearch||Window|||49|
31|If Image/Pixel Found||1|0|If_Statement|||||51|
032|[Pause]||1|1000|Sleep|||||53|
33|Method:Click:TagName||1|0|IECOM_Set|INPUT:6||||55|
034|[Pause]||1|1000|Sleep|||||57|
35|Method:Click:TagName||1|0|IECOM_Set|INPUT:7||||59|
36|[End If]|EndIf|1|0|If_Statement|||||61|
37|[Pause]||1|5000|Sleep|||||62|
38|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200916211549.png|1|0|ImageSearch|UntilFound|Window|||63|
39|[Pause]||1|2000|Sleep|||||72|
40|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png|1|0|ImageSearch|UntilFound|Window|||73|
41|[Pause]||1|2000|Sleep|||||82|
42|Continue, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201001114340.png|1|0|ImageSearch|UntilFound|Window|||83|
43|If Image/Pixel Found||1|0|If_Statement|||||90|
44|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||92|
45|[End If]|EndIf|1|0|If_Statement|||||94|
046|[MsgBox]|%codhtml%|1|0|MsgBox|0||||95|
47|[Pause]||1|4000|Sleep|||||97|
48|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200921213437.png|1|0|ImageSearch|UntilFound|Window|||99|
49|[Pause]||1|2000|Sleep|||||108|
50|End|{End}|1|0|Send|||||109|
51|[Pause]||1|2000|Sleep|||||110|
52|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200921213703.png|1|0|ImageSearch||Window|||111|
53|If Image/Pixel Not Found||1|0|If_Statement|||||116|
54|[Pause]||1|2000|Sleep|||||118|
55|End|{End}|1|0|Send|||||119|
56|[Pause]||1|2000|Sleep|||||120|
57|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200921213703.png|1|0|ImageSearch||Window|||121|
58|[End If]|EndIf|1|0|If_Statement|||||126|
59|[Pause]||1|1000|Sleep|||||127|

[PMC Code v5.3.4]|||1|Window,2,Fast,0,1,Input,-1,-1,1|1|auxSarch()
Context=None||None|
Groups=Start:1
1|[FuncParameter]|cuenta|1|0|FuncParameter|||||1|
2|[FuncParameter]|beneficiario|1|0|FuncParameter|||||1|
3|[FuncParameter]|cc|1|0|FuncParameter|||||1|
4|[FuncParameter]|valor|1|0|FuncParameter|||||1|
5|[FuncParameter]|referencia|1|0|FuncParameter|||||1|
6|[FuncParameter]|img|1|0|FuncParameter|||||1|
7|[FuncParameter]|page|1|0|FuncParameter|||||1|
8|[FuncParameter]|pathImg|1|0|FuncParameter|||||1|
9|[FuncParameter]|pathDescargas|1|0|FuncParameter|||||1|
10|[FuncParameter]|pathComprobantes|1|0|FuncParameter|||||1|
11|[FunctionStart]|auxSarch|1|0|UserFunction|Local| / |||1|
12|[Pause]||1|500|Sleep|||||3|
13|[MsgBox]|estoy en AuxSearch y la cedula es: %cc%|1|0|MsgBox|0||||4|
14|searchDynamics|comprobante := cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,pathComprobantes|1|0|Function|||||5|
15|[Pause]||1|3000|Sleep|||||6|
016|[MsgBox]|%comprobante%|1|0|MsgBox|0||||7|
17|Compare Variables|comprobante == 1|1|0|If_Statement|||||9|
18|FileAppend|`n  nombre:%beneficiario%.Documento:%cc%``,Valor:%valor%``,pagina:%page%hora:%A_Tab%%A_Hour%:%A_Min%:%A_Sec%%A_Tab%, %pathComprobantes%enviados.txt|1|0|FileAppend|||||12|
19|[Pause]||1|300|Sleep|||||13|
20|[Else]|Else|1|0|If_Statement|||||14|
21|searchDynamics|auxComprobante := cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,pathComprobantes|1|0|Function|||||17|
22|[Pause]||1|3000|Sleep|||||18|
23|Compare Variables|auxComprobante == 0|1|0|If_Statement|||||19|
24|FileAppend|`n  nombre:%beneficiario%.Documento:%cc%``,Valor:%valor%``,pagina:%page%hora:%A_Tab%%A_Hour%:%A_Min%:%A_Sec%%A_Tab%, %pathComprobantes%inconvenientes.txt|1|0|FileAppend|||||21|
25|[Else]|Else|1|0|If_Statement|||||22|
26|FileAppend|`n  nombre:%beneficiario%.Documento:%cc%``,Valor:%valor%``,pagina:%page%hora:%A_Tab%%A_Hour%:%A_Min%:%A_Sec%%A_Tab%, %pathComprobantes%enviados.txt|1|0|FileAppend|||||25|
27|[End If]|EndIf|1|0|If_Statement|||||26|
28|[End If]|EndIf|1|0|If_Statement|||||27|
29|[FuncReturn]|comprobante|1|0|FuncReturn|||||28|

[PMC Code v5.3.4]|||1|Window,2,Fast,0,1,Input,-1,-1,1|1|auxAbonado()
Context=None||None|
Groups=Start:1
1|[FuncParameter]|cuenta|1|0|FuncParameter|||||1|
2|[FuncParameter]|beneficiario|1|0|FuncParameter|||||1|
3|[FuncParameter]|cc|1|0|FuncParameter|||||1|
4|[FuncParameter]|valor|1|0|FuncParameter|||||1|
5|[FuncParameter]|referencia|1|0|FuncParameter|||||1|
6|[FuncParameter]|img|1|0|FuncParameter|||||1|
7|[FuncParameter]|page|1|0|FuncParameter|||||1|
8|[FuncParameter]|pathImg|1|0|FuncParameter|||||1|
9|[FuncParameter]|pathDescargas|1|0|FuncParameter|||||1|
10|[FuncParameter]|abonados|1|0|FuncParameter|||||1|
11|[FuncParameter]|pathComprobantes|1|0|FuncParameter|||||1|
12|[FunctionStart]|auxAbonado|1|0|UserFunction|Local| / |||1|
13|[MsgBox]|%cc% en auxAbonado|1|0|MsgBox|0||||3|
14|Compare Variables|abonados == 1|1|0|If_Statement|||||4|
015|[MsgBox]|Lugar de llamada a la funcion de busqueda de cedula y envio de correo|1|0|MsgBox|0||||6|
16|[Pause]||1|1000|Sleep|||||8|
17|auxSarch|comprobante := cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,pathComprobantes|1|0|Function|||||10|
18|[Pause]||1|1000|Sleep|||||11|
19|[Else]|Else|1|0|If_Statement|||||12|
20|FileAppend|`n  nombre:%beneficiario%.Documento:%cc%``,Valor:%valor%``,pagina:%page%hora:%A_Tab%%A_Hour%:%A_Min%:%A_Sec%%A_Tab%, %pathComprobantes%inconvenientes.txt|1|0|FileAppend|||||15|
21|[End If]|EndIf|1|0|If_Statement|||||16|
022|Push|_null := % resultArray[%A_Index%+1]|1|0|Method|imagen||||17|
023|Push|_null := resultArray[A_Index+1]|1|0|Method|imagen||||19|
024|[Assign Variable]|banImg := 1|1|0|Variable|Expression||||20|
025|[ElseIf] Compare Variables|banImg = 1|1|0|If_Statement|||||21|
026|Continue, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201008134856.png|1|0|ImageSearch||Window|||24|
027|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||27|
028|Continue, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201008135352.png|1|0|ImageSearch||Window|||29|
029|[Pause]||1|500|Sleep|||||32|
030|If Image/Pixel Found||1|0|If_Statement|||||33|
031|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||35|
032|[End If]|EndIf|1|0|If_Statement|||||37|
033|[Pause]||1|1000|Sleep|||||38|
034|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png|1|0|ImageSearch|UntilFound|Window|||39|
035|[Pause]||1|1500|Sleep|||||48|
036|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201007154153.png|1|0|ImageSearch||Window|||49|
037|[Pause]||1|1000|Sleep|||||54|
038|If Image/Pixel Found||1|0|If_Statement|UntilFound||||55|
039|Enter|{Enter}|1|0|Send|||||57|
040|[Pause]||1|500|Sleep|||||58|
041|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201007151603.png|1|0|ImageSearch|UntilFound|Window|||59|
042|[Pause]||1|500|Sleep|||||68|
043|[Assign Variable]|Clipboard := Resuelto.|1|0|Variable|||||69|
044|[Pause]||1|300|Sleep|||||70|
045|Control + v|{Control Down}{v}{Control Up}|1|0|Send|||||71|
046|[Pause]||1|500|Sleep|||||72|
047|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20200924082550.png|1|0|ImageSearch|UntilFound|Window|||73|
048|[Pause]||1|2000|Sleep|||||82|
049|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201007151912.png|1|0|ImageSearch|UntilFound|Window|||83|
050|[Pause]||1|3000|Sleep|||||92|
051|[Else]|Else|1|0|If_Statement|||||93|
052|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20200924082550.png|1|0|ImageSearch|UntilFound|Window|||96|
053|[End If]|EndIf|1|0|If_Statement|||||105|
054|[Pause]||1|2000|Sleep|||||106|
055|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200929114218.png|1|0|ImageSearch|UntilFound|Window|||107|
056|[Pause]||1|2000|Sleep|||||116|
057|If Window Active|Mensaje de página web|1|0|If_Statement|Button1||||117|
058|Left Click|Left, 1,  NA|1|10|ControlClick|Button1|Mensaje de página web|||119|
059|[Pause]||1|1000|Sleep|||||121|
060|[End If]|EndIf|1|0|If_Statement|||||122|
061|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||123|
062|[Pause]||1|1000|Sleep|||||125|
063|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png|1|0|ImageSearch|UntilFound|Window|||126|
064|[Pause]||1|2000|Sleep|||||135|
065|[End If]|EndIf|1|0|If_Statement|||||136|
066|[Else]|Else|1|0|If_Statement|||||137|
067|[Assign Variable]|status := 0|1|0|Variable|||||140|
068|[Pause]||1|3000|Sleep|||||141|
069|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20200923182420.png|1|0|ImageSearch|UntilFound|Window|||142|
070|[Pause]||1|10000|Sleep|||||151|
071|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200924143844.png|1|0|ImageSearch|UntilFound|Window|||152|
072|[Pause]||1|1000|Sleep|||||161|
073|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png|1|0|ImageSearch|UntilFound|Window|||162|
074|[Pause]||1|1000|Sleep|||||171|
075|[End If]|EndIf|1|0|If_Statement|||||172|
076|[FuncReturn]|status|1|0|FuncReturn|||||173|

[PMC Code v5.3.4]|||1|Window,2,Fast,0,1,Input,-1,-1,1|1|searchDynamics()
Context=None||None|
Groups=Iniciar:1
1|[FuncParameter]|cuenta|1|0|FuncParameter|||||1|
2|[FuncParameter]|beneficiario|1|0|FuncParameter|||||1|
3|[FuncParameter]|cc|1|0|FuncParameter|||||1|
4|[FuncParameter]|valor|1|0|FuncParameter|||||1|
5|[FuncParameter]|referencia|1|0|FuncParameter|||||1|
6|[FuncParameter]|img|1|0|FuncParameter|||||1|
7|[FuncParameter]|page|1|0|FuncParameter|||||1|
8|[FuncParameter]|pathImg|1|0|FuncParameter|||||1|
9|[FuncParameter]|pathDescargas|1|0|FuncParameter|||||1|
10|[FuncParameter]|pathComprobantes|1|0|FuncParameter|||||1|
11|[FunctionStart]|searchDynamics|1|0|UserFunction|Global| / |||1|
12|[MsgBox]| estoy en searchDinamycs y esta es la cedula %cc%|1|0|MsgBox|0||||4|
13|[Assign Variable]|status := 1|1|0|Variable|||||5|
14|[Pause]||1|6000|Sleep|||||6|
15|Continue, Continue, FoundX, FoundY, 0|0, 0, 1280, 1024, %pathImg%Screen_20201022080131.png|1|0|ImageSearch||Screen|||7|
16|If Image/Pixel Not Found||1|0|If_Statement|||||9|
17|[Pause]||1|2000|Sleep|||||11|
18|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022080228.png|1|0|ImageSearch|UntilFound|Window|||12|
19|End|{End}|1|0|Send|||||21|
20|searchDynamics|auxComprobante := cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,pathComprobantes|1|0|Function|||||22|
21|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022095157.png|1|0|ImageSearch|UntilFound|Window|||23|
22|[End If]|EndIf|1|0|If_Statement|||||32|
23|[Pause]||1|3000|Sleep|||||33|
024|[MsgBox]|Antes de Maximizar|1|0|MsgBox|0||||34|
025|WinMaximize||1|333|WinMaximize||Casos Vista de devoluciones de dinero - Microsoft Dynamics 365 - Internet Explorer|||36|
26|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022151658.png|1|0|ImageSearch|UntilFound|Window|||38|
27|[Pause]||1|5000|Sleep|||||48|
28|Continue, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%filterWhite.PNG|1|0|ImageSearch||Window|||49|
29|If Image/Pixel Found||1|0|If_Statement|UntilFound||||52|
30|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||54|
31|[Pause]||1|2000|Sleep|||||56|
32|Move|0, 0, 0|1|10|Click|||||57|
33|[End If]|EndIf|1|0|If_Statement|||||59|
034|[MsgBox]|buscando filtro negro|1|0|MsgBox|0||||60|
35|[Pause]||1|1000|Sleep|||||62|
36|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%blackFilter.PNG|1|0|ImageSearch|UntilFound|Window|||64|
037|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022080608.png|1|0|ImageSearch|UntilFound|Window|||73|
38|[Pause]||1|2000|Sleep|||||83|
39|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022080712.png|1|0|ImageSearch|UntilFound|Window|||85|
40|[Pause]||1|500|Sleep|||||94|
41|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022080759.png|1|0|ImageSearch|UntilFound|Window|||95|
42|[Pause]||1|500|Sleep|||||104|
43|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%aceptarDD.PNG|1|0|ImageSearch|UntilFound|Window|||105|
44|[Pause]||1|2000|Sleep|||||114|
45|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022080849.png|3|0|ImageSearch|UntilFound|Window|||115|
46|[Pause]||1|3000|Sleep|||||124|
47|Tab|{Tab}|1|0|Send|||||125|
48|[Pause]||1|1000|Sleep|||||126|
49|Enter|{Enter}|1|0|Send|||||127|
50|[Pause]||1|5000|Sleep|||||128|
51|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201022080916.png|1|0|ImageSearch|UntilFound|Window|||129|
52|[Pause]||1|1000|Sleep|||||138|
53|[LoopStart]|LoopStart|5|0|Loop|||||139|
54|Right|{Right}|1|0|Send|||||141|
55|[Pause]||1|1000|Sleep|||||142|
56|[LoopEnd]|LoopEnd|1|0|Loop|||||143|
57|[Assign Variable]|startX := 0|1|0|Variable|||||144|
58|[LoopStart]|LoopStart|7|0|Loop|||||145|
59|Continue, Continue, FoundX, FoundY, 1|%startX%, 0, 1920, 1080, %pathImg%Screen_20201022080940.png|3|0|ImageSearch|UntilFound|Window|||147|
60|[Assign Variable]|startX := %FoundX%|1|0|Variable|||||154|
61|[LoopEnd]|LoopEnd|1|0|Loop|||||155|
62|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||156|
63|[Pause]||1|300|Sleep|||||158|
64|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022081032.png|1|0|ImageSearch|UntilFound|Window|||159|
65|[Pause]||1|2000|Sleep|||||168|
66|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022081251.png|1|0|ImageSearch|UntilFound|Window|||169|
67|[Pause]||1|2000|Sleep|||||178|
68|Continue, Continue, FoundX, FoundY, 0|0, 0, 1280, 1024, %pathImg%selectcc.PNG|1|0|ImageSearch||Window|||179|
69|[Pause]||1|1000|Sleep|||||181|
70|If Image/Pixel Found||1|0|If_Statement|||||182|
71|[Assign Variable]|bccfiltro := 1|1|0|Variable|||||184|
72|[Pause]||1|1000|Sleep|||||185|
73|[End If]|EndIf|1|0|If_Statement|||||186|
74|Down|{Down}|1|0|Send|||||187|
75|[Pause]||1|500|Sleep|||||188|
76|Enter|{Enter}|1|0|Send|||||189|
77|[Pause]||1|500|Sleep|||||190|
78|Tab|{Tab}|1|0|Send|||||191|
79|[Pause]||1|1000|Sleep|||||192|
80|[Assign Variable]|Clipboard := %cc%|1|0|Variable|||||193|
81|[Pause]||1|1000|Sleep|||||194|
082|[MsgBox]|%cc%`, Esta es la cedula|1|0|MsgBox|0||||195|
83|Control + v|{Control Down}{v}{Control Up}|1|0|Send|||||197|
84|[Pause]||1|1000|Sleep|||||199|
85|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%aceptarCC.PNG|1|0|ImageSearch|UntilFound|Window|||200|
86|[Pause]||1|4000|Sleep|||||209|
087|[MsgBox]|Buscanco "Bu"|1|0|MsgBox|0||||210|
88|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201103103155.png|1|0|ImageSearch|UntilFound|Window|||212|
89|If Image/Pixel Not Found||1|0|If_Statement|||||222|
90|[Pause]||1|1000|Sleep|||||224|
091|[MsgBox]|Buscando Esquina|1|0|MsgBox|0||||225|
92|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022082733.png|1|0|ImageSearch||Window|||227|
93|[End If]|EndIf|1|0|If_Statement|||||233|
94|[Pause]||1|2000|Sleep|||||234|
95|Control + f|{Control Down}{f}{Control Up}|1|0|Send|||||235|
96|[Pause]||1|2000|Sleep|||||236|
97|[Assign Variable]|Clipboard := %valor%|1|0|Variable|||||237|
98|[Pause]||1|1000|Sleep|||||238|
99|Control + v|{Control Down}{v}{Control Up}|1|0|Send|||||239|
100|[Pause]||1|1000|Sleep|||||240|
101|Enter|{Enter}|1|0|Send|||||241|
102|[Pause]||1|2000|Sleep|||||242|
0103|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022083430.png|1|0|ImageSearch||Window|||243|
0104|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%encuentroValor.PNG|1|0|ImageSearch||Window|||249|
105|Left Click, Continue, FoundX, FoundY, 0|0, 0, 1280, 1024, %pathImg%findValue.PNG|1|0|ImageSearch||Window|||254|
106|[Pause]||1|200|Sleep|||||259|
0107|[MsgBox]|%bccfiltro%`,%ErrorLevel%|1|0|MsgBox|0||||260|
108|Evaluate Expression|bccfiltro==1 && ErrorLevel==0|1|0|If_Statement|UntilFound||||262|
109|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||265|
110|[Pause]||1|10000|Sleep|||||267|
111|Control + f|{Control Down}{f}{Control Up}|1|0|Send|||||268|
112|[Pause]||1|2000|Sleep|||||269|
113|Continue, Continue, FoundX, FoundY|0, 0, 1280, 1024, 0x0078D7, 0, Fast RGB|1|0|PixelSearch||Window|||270|
114|If Image/Pixel Not Found||1|0|If_Statement|||||272|
115|Control + f|{Control Down}{f}{Control Up}|1|0|Send|||||274|
116|[Pause]||1|2000|Sleep|||||275|
117|[End If]|EndIf|1|0|If_Statement|||||276|
118|[Assign Variable]|Clipboard := Agregar Llama|1|0|Variable|||||277|
119|[Pause]||1|2000|Sleep|||||278|
120|Control + v|{Control Down}{v}{Control Up}|1|0|Send|||||279|
121|[Pause]||1|2000|Sleep|||||280|
122|[Assign Variable]|startY := 0|1|0|Variable|||||281|
123|[LoopStart]|LoopStart|2|0|Loop|||||282|
124|Continue, Continue, FoundX, FoundY, 1|0, %startY%, 1920, 1080, %pathImg%Screen_20201022083533.png|1|0|ImageSearch||Window|||284|
125|[Assign Variable]|startY := %FoundY%|1|0|Variable|||||287|
126|[LoopEnd]|LoopEnd|1|0|Loop|||||288|
127|[Pause]||1|500|Sleep|||||289|
128|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||290|
0129|[MsgBox]|%ErrorLevel%|1|0|MsgBox|0||||292|
130|If Image/Pixel Not Found||1|0|If_Statement|||||294|
131|[Assign Variable]|status := 0|1|0|Variable|||||297|
132|[Pause]||1|3000|Sleep|||||298|
133|Method:Navigate:|https://gco.crm2.dynamics.com/main.aspx|1|0|IECOM_Set|:|LoadWait|||299|
0134|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201022083610.png|1|0|ImageSearch|UntilFound|Window|||304|
135|[Pause]||1|10000|Sleep|||||314|
136|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022094907.png|1|0|ImageSearch|UntilFound|Window|||316|
137|[Pause]||1|1000|Sleep|||||325|
138|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022083735.png|1|0|ImageSearch|UntilFound|Window|||326|
139|[End If]|EndIf|1|0|If_Statement|||||335|
140|Evaluate Expression|%status%==1|1|0|If_Statement|UntilFound||||336|
141|[Pause]||1|1000|Sleep|||||338|
142|Enter|{Enter}|1|0|Send|||||339|
143|[Pause]||1|1000|Sleep|||||340|
144|[Assign Variable]|Clipboard := ""|1|0|Variable|||||341|
145|[Pause]||1|2000|Sleep|||||342|
0146|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022083826.png|1|0|ImageSearch|UntilFound|Window|||343|
0147|[Pause]||1|2000|Sleep|||||353|
0148|Control + c|{Control Down}{c}{Control Up}|2|0|Send|||||354|
0149|[Pause]||1|2000|Sleep|||||355|
0150|Control + c|{Control Down}{c}{Control Up}|1|0|Send|||||356|
0151|[Pause]||1|2000|Sleep|||||357|
0152|If Window Active|ahk_class #32770|1|0|If_Statement|Button2||||358|
0153|Left Click|Left, 1,  NA|1|10|ControlClick|Button2|ahk_class #32770|||360|
0154|[Pause]||1|2000|Sleep|||||362|
0155|[End If]|EndIf|1|0|If_Statement|||||363|
0156|Compare Variables|Clipboard = ""|1|0|If_Statement|||||364|
157|[Assign Variable]|Clipboard := Radicado Devolucion de Dinero|1|0|Variable|||||366|
0158|[MsgBox]|%Clipboard%|1|0|MsgBox|0||||368|
0159|[End If]|EndIf|1|0|If_Statement|||||370|
160|[Pause]||1|500|Sleep|||||371|
161|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022083846.png|1|0|ImageSearch|UntilFound|Window|||373|
162|[Pause]||1|2000|Sleep|||||382|
163|Control + v|{Control Down}{v}{Control Up}|1|0|Send|||||383|
164|[Pause]||1|3000|Sleep|||||384|
165|[LoopStart]|LoopStart|7|0|Loop|||||385|
166|Tab|{Tab}|1|0|Send|||||387|
167|[Pause]||1|100|Sleep|||||388|
168|[LoopEnd]|LoopEnd|1|0|Loop|||||389|
169|[Assign Variable]|Clipboard := Hola, `n`nQueremos contarte que la devolución del dinero ya fue efectiva, te enviamos el comprobante de transferencia.`n`nCualquier inquietud adicional, con gusto será atendida.`n`nSERVICIO AL CLIENTE.|1|0|Variable|||||390|
170|Control + v|{Control Down}{v}{Control Up}|1|0|Send|||||400|
171|[Pause]||1|500|Sleep|||||401|
0172|[MsgBox]|%pathImg%`,%pathDescargas%|1|0|MsgBox|0||||402|
173|b64Jpg|_null := img,pathImg,pathDescargas|1|0|Function|||||404|
174|[Pause]||1|1000|Sleep|||||406|
175|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022083904.png|1|0|ImageSearch|UntilFound|Window|||407|
176|[Pause]||1|2000|Sleep|||||416|
177|Alt + Tab|{Alt Down}{Tab}{Alt Up}|1|0|Send|||||417|
178|[Pause]||1|500|Sleep|||||418|
179|Alt + Tab|{Alt Down}{Tab}{Alt Up}|1|0|Send|||||419|
180|[Pause]||1|500|Sleep|||||420|
181|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022085611.png|1|0|ImageSearch|UntilFound|Window|||421|
182|[Pause]||1|2000|Sleep|||||430|
183|[Control]|%pathDescargas%test-download.png|1|0|ControlSetText|Edit1|ahk_class #32770|||431|
184|[Pause]||1|4000|Sleep|||||432|
185|Left Click|Left, 1,  NA|1|10|ControlClick|Button1|ahk_class #32770|||433|
186|[Pause]||1|1000|Sleep|||||435|
187|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022113123.png|1|0|ImageSearch|UntilFound|Window|||436|
188|[Pause]||1|3000|Sleep|||||445|
189|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20201022113437.png|1|0|ImageSearch|UntilFound|Window|||446|
190|[Pause]||1|4000|Sleep|||||455|
0191|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201022084057.png|1|0|ImageSearch|UntilFound|Window|||456|
0192|[Pause]||1|6000|Sleep|||||466|
0193|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20200924033234.png|1|0|ImageSearch||Window|||467|
0194|[Pause]||1|2000|Sleep|||||472|
0195|Continue, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201008135352.png|1|0|ImageSearch||Window|||473|
0196|[Pause]||1|500|Sleep|||||476|
0197|If Image/Pixel Found||1|0|If_Statement|||||477|
0198|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||479|
0199|[End If]|EndIf|1|0|If_Statement|||||481|
0200|[Pause]||1|1000|Sleep|||||482|
0201|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png|1|0|ImageSearch|UntilFound|Window|||483|
0202|[Pause]||1|1500|Sleep|||||492|
0203|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201007154153.png|1|0|ImageSearch||Window|||493|
0204|[Pause]||1|1000|Sleep|||||498|
0205|If Image/Pixel Found||1|0|If_Statement|UntilFound||||499|
0206|Enter|{Enter}|1|0|Send|||||501|
0207|[Pause]||1|500|Sleep|||||502|
0208|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201007151603.png|1|0|ImageSearch|UntilFound|Window|||503|
0209|[Pause]||1|500|Sleep|||||512|
0210|[Assign Variable]|Clipboard := Resuelto.|1|0|Variable|||||513|
0211|[Pause]||1|300|Sleep|||||514|
0212|Control + v|{Control Down}{v}{Control Up}|1|0|Send|||||515|
0213|[Pause]||1|500|Sleep|||||516|
0214|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20200924082550.png|1|0|ImageSearch|UntilFound|Window|||517|
0215|[Pause]||1|2000|Sleep|||||526|
0216|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20201007151912.png|1|0|ImageSearch|UntilFound|Window|||527|
0217|[Pause]||1|3000|Sleep|||||536|
0218|[Else]|Else|1|0|If_Statement|||||537|
0219|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20200924082550.png|1|0|ImageSearch|UntilFound|Window|||540|
0220|[End If]|EndIf|1|0|If_Statement|||||549|
0221|[Pause]||1|2000|Sleep|||||550|
0222|[MsgBox]|inicia Buscque de imagen X|1|0|MsgBox|0||||551|
223|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%xmessage.PNG|1|0|ImageSearch|UntilFound|Window|||552|
224|[Pause]||1|2000|Sleep|||||562|
0225|If Window Active|Mensaje de página web|1|0|If_Statement|Button1||||563|
226|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%aceptarMesX.PNG|1|0|ImageSearch|UntilFound|Window|||566|
0227|Left Click|Left, 1,  NA|1|10|ControlClick|Button1|Mensaje de página web|||576|
228|[Pause]||1|1000|Sleep|||||579|
0229|[End If]|EndIf|1|0|If_Statement|||||581|
230|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%xmessage.PNG|1|0|ImageSearch|UntilFound|Window|||583|
0231|Left Move & Click|%FoundX%, %FoundY% Left, 1|1|10|Click|||||593|
232|[Pause]||1|1000|Sleep|||||596|
0233|[MsgBox]|a continuacion se minimiza pestaña|1|0|MsgBox|0||||598|
234|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png|1|0|ImageSearch|UntilFound|Window|||600|
235|[Pause]||1|2000|Sleep|||||610|
236|[End If]|EndIf|1|0|If_Statement|||||611|
237|[Else]|Else|1|0|If_Statement|||||612|
238|[Assign Variable]|status := 0|1|0|Variable|||||615|
239|[Pause]||1|3000|Sleep|||||616|
240|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1280, 1024, %pathImg%Screen_20200923182420.png|1|0|ImageSearch|UntilFound|Window|||617|
241|[Pause]||1|10000|Sleep|||||626|
242|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200924143844.png|1|0|ImageSearch|UntilFound|Window|||627|
243|[Pause]||1|1000|Sleep|||||636|
244|Left Click, Continue, FoundX, FoundY, 1|0, 0, 1920, 1080, %pathImg%Screen_20200918143624.png|1|0|ImageSearch|UntilFound|Window|||637|
245|[Pause]||1|1000|Sleep|||||646|
246|[End If]|EndIf|1|0|If_Statement|||||647|
247|[FuncReturn]|status|1|0|FuncReturn|||||648|

*/
