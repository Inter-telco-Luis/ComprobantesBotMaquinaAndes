; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1


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
Goto, MacroNro7
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
    /*
    MsgBox, 0, , estoy en AuxSearch y la cedula es: %cc%
    */
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
    /*
    MsgBox, 0, , %cc% en auxAbonado
    */
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
    /*
    MsgBox, 0, ,  estoy en searchDinamycs y esta es la cedula %cc%
    */
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
        Sleep, 1000
        Send, {End}
        Sleep, 1000
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022095157.png
            CenterImgSrchCoords(pathImg "Screen_20201022095157.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
        Sleep, 1000
        auxComprobante := searchDynamics(cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,pathComprobantes)
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
        Click, 400, 0, 0
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
    Loop, 5
    {
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022080712.png
        CenterImgSrchCoords(pathImg "Screen_20201022080712.png", FoundX, FoundY)
        If ErrorLevel = 0
        	Click, %FoundX%, %FoundY% Left, 1
        Sleep, 500
        If (ErrorLevel = 0)
        {
            Break
        }
        Click, %FoundX%, %FoundY% Left, 1
        Sleep, 10
        Sleep, 500
    }
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
    Sleep, 1000
    IfWinActive, Mensaje de la página web
    {
        Send, {Enter}
        Sleep, 300
        bccfiltro := 0
        Sleep, 1000
        Send, {Tab}
        Sleep, 1000
        Send, {Enter}
    }
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
        Sleep, 12000
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
        Sleep, 1000
        statusOne(status,referencia,img,pathImg,pathDescargas)
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

MacroNro7:
abonados := 1
cuenta := ""
beneficiario := ""
cc := ""
valor := ""
img := 0
referencia := ""
page := ""
resultArray := []
flagNext := 0
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
Sleep, 1000
loginDynamics(pathImg)
Sleep, 500
For key, value in resultArray
{
    Sleep, 200
    /*
    MsgBox, 0, , %value%`,%key%`,%flagNext%
    */
    If (value == "No_Abonados")
    {
        abonados := 0
    }
    Else If (value == "numero de cuenta")
    {
        flagNext := 1
    }
    Else If (flagNext=="1")
    {
        cuenta := value
        flagNext := 0
    }
    Else If (value == "nombre de beneficiario")
    {
        flagNext := 2
    }
    Else If (flagNext=="2")
    {
        beneficiario := value
        flagNext := 0
    }
    Else If (value == "documento")
    {
        flagNext := 3
    }
    Else If (flagNext=="3")
    {
        cc := value
        flagNext := 0
    }
    Else If (value == "valor")
    {
        flagNext := 4
    }
    Else If (flagNext=="4")
    {
        valor := value
        valor .= "."
        flagNext := 5
    }
    Else If (flagNext=="5")
    {
        valor .= value
        flagNext := 0
    }
    Else If (value == "referencia")
    {
        flagNext := 6
    }
    Else If (flagNext=="6")
    {
        referencia := value
        flagNext := 0
    }
    Else If (value == "img")
    {
        flagNext := 7
    }
    Else If (flagNext=="7")
    {
        img := value
        flagNext := 0
    }
    Else If (value == "page")
    {
        flagNext := 8
    }
    Else If (flagNext=="8")
    {
        page := value
        Sleep, 1000
        /*
        MsgBox, 0, , %cc%`,estoy en Macro7
        */
        auxAbonado(cuenta,beneficiario,cc,valor,referencia,img,page,pathImg,pathDescargas,abonados,pathComprobantes)
        flagNext := 0
        Sleep, 1000
    }
}
Return

statusOne(status, referencia, img, pathImg, pathDescargas)
{
    If (status==1)
    {
        Sleep, 1000
        Send, {Enter}
        Sleep, 1000
        Clipboard := ""
        Sleep, 2000
        Clipboard := "Radicado Devolucion de Dinero"
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
        Sleep, 1000
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%Screen_20201022085611.png
            CenterImgSrchCoords(pathImg "Screen_20201022085611.png", FoundX, FoundY)
            If ErrorLevel = 0
            	Click, %FoundX%, %FoundY% Left, 1
        }
        Until ErrorLevel = 0
        Sleep, 500
        Click, 400, 0, 0
        Sleep, 10
        Sleep, 1000
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
        ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, C:\img\confirmarRC0.PNG
        CenterImgSrchCoords("C:\img\confirmarRC0.PNG", FoundX, FoundY)
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
            Sleep, 1000
            Loop
            {
                CoordMode, Pixel, Window
                ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, C:\img\Resolucion2.PNG
                CenterImgSrchCoords("C:\img\Resolucion2.PNG", FoundX, FoundY)
                If ErrorLevel = 0
                	Click, %FoundX%, %FoundY% Left, 1
            }
            Until ErrorLevel = 0
            Sleep, 1000
            Clipboard := "Resuelto."
            Sleep, 1000
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
                ImageSearch, FoundX, FoundY, 0, 0, 1280, 1024, C:\img\ResolverZ.PNG
                CenterImgSrchCoords("C:\img\ResolverZ.PNG", FoundX, FoundY)
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
        Sleep, 1000
        Loop
        {
            CoordMode, Pixel, Window
            ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, %pathImg%xmessage.PNG
            CenterImgSrchCoords(pathImg "xmessage.PNG", FoundX, FoundY)
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
        Sleep, 2000
    }
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
