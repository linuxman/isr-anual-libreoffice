REM  *****  BASIC  *****
 
Option Explicit
 
Function ISR_Anual_2021(ByVal PercepcionesGravables As Double) As Double
 
'*****************************************************************************************************
' FUNCION PARA CALCULAR EL ISR ANUAL
' Copyright (C) Francisco Javier de la Torre '
' Este código es software libre. Puede redistribuirlo y/o modificarlo bajo los términos de la
' Licencia Pública General de GNU según es publicada por la Free Software Foundation, bien de la
' versión 2 de dicha Licencia o bien (según su elección) de cualquier versión posterior.
' Este programa se distribuye con la esperanza de que sea útil, pero SIN NINGUNA GARANTÍA, incluso
' sin la garantía MERCANTIL implícita o sin garantizar la CONVENIENCIA PARA UN PROPÓSITO PARTICULAR.
' Véase la Licencia Pública General de GNU para más detalles.
' Debería haber recibido una copia de la Licencia Pública General junto con este programa. Si no ha
' sido así, escriba a la Free Software Foundation, Inc., en 675 Mass Ave, Cambridge, MA 02139, EEUU.
'
' LinuxmanR4
' https://linuxmanr4.com
' 2021
'*****************************************************************************************************
    Dim ISR_anual(11, 2) As Double
    Dim SUBSIDIO_AL_EMPLEO_ANUAL(11, 1) As Double
    Dim ISR_LimiteInferior As Double
    Dim CuotaFija As Double
    Dim PorcentajeSobreExcedente As Double
    Dim i As Integer
    Dim ISR As Double
 
    'Definición de las tablas iniciales
 
    'ISR ANUAL
    '==============================
    'Limite inferior
 
    ISR_anual(0, 0) = 0.01
    ISR_anual(1, 0) = 7735.00
    ISR_anual(2, 0) = 65651.07
    ISR_anual(3, 0) = 115375.90
    ISR_anual(4, 0) = 134119.41
    ISR_anual(5, 0) = 160577.65
    ISR_anual(6, 0) = 323862.00
    ISR_anual(7, 0) = 510451.00
    ISR_anual(8, 0) = 974535.03
    ISR_anual(9, 0) = 1299380.04
    ISR_anual(10, 0) = 3898140.12
    ISR_anual(11, 0) = 999999999   'Limite superior muy alto
 
    'Cuota fija
    ISR_anual(0, 1) = 0
    ISR_anual(1, 1) = 148.51
    ISR_anual(2, 1) = 3855.14
    ISR_anual(3, 1) = 9265.20
    ISR_anual(4, 1) = 12264.16
    ISR_anual(5, 1) = 17005.47
    ISR_anual(6, 1) = 51883.01
    ISR_anual(7, 1) = 95768.74
    ISR_anual(8, 1) = 234993.95
    ISR_anual(9, 1) = 338944.34
    ISR_anual(10, 1) = 1222522.76
    ISR_anual(11, 1) = 0
 
    'Porcentaje sobre excedente
    ISR_anual(0, 2) = 0.0192
    ISR_anual(1, 2) = 0.064
    ISR_anual(2, 2) = 0.1088
    ISR_anual(3, 2) = 0.16
    ISR_anual(4, 2) = 0.1792
    ISR_anual(5, 2) = 0.2136
    ISR_anual(6, 2) = 0.2352
    ISR_anual(7, 2) = 0.3
    ISR_anual(8, 2) = 0.32
    ISR_anual(9, 2) = 0.34
    ISR_anual(10, 2) = 0.35
    ISR_anual(11, 2) = 0
 
    'Iniciamos el cálculo del ISR anual.
 
    CuotaFija = 0: PorcentajeSobreExcedente = 0
    'Buscamos un valor apropiado en la tabla del ISR Anual
    i = 0
    Do
        If ISR_anual(i, 0) > PercepcionesGravables Then
            ISR_LimiteInferior = ISR_anual(i - 1, 0)
            CuotaFija = ISR_anual(i - 1, 1)
            PorcentajeSobreExcedente = ISR_anual(i - 1, 2)
            Exit Do
        Else
            i = i + 1
        End If
    Loop Until i = 12
 
    'Ya tenemos los valores de Cuota Fija y Porcentaje sobre excedente, procedemos a calcular el ISR Anual
 
    ISR = CuotaFija + ((PercepcionesGravables - ISR_LimiteInferior) * PorcentajeSobreExcedente)
 
    ISR_anual_2021 = Format(ISR,"000000000000000.00")
 
End Function
