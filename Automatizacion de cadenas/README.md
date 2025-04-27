**Macro de Automatización para Hoja1 - Descripción y Uso**

Este documento explica el comportamiento y la lógica de la macro VBA asociada al evento `Worksheet_Change` en la hoja llamada **Hoja1**. La macro implementa varias funcionalidades para automatizar tareas comunes de validación, escritura de datos y formateo.

---

## Índice

1. [Requisitos](#requisitos)
1. [Instalación](#instalaci%C3%B3n)
1. [Resumen de Funcionalidades](#resumen-de-funcionalidades)
1. [Detalle de Módulos y Subs](#detalle-de-m%C3%B3dulos-y-subs)
   - [1. `escribirAutomaticamente`](#1-escribierautomaticamente)
   - [2. `RevertValueInF`](#2-revertvalueinf)
   - [3. `Worksheet_Change`](#3-worksheet_change)
     - Copia de B4 a C4
     - Clasificación en columna P
     - Formato numérico en columna G
     - Procesamiento de textos en columna J
     - Resaltado en columna I

---

## Requisitos

- Microsoft Excel (versión que soporte VBA).
- Hoja de cálculo llamada **Hoja1** en el libro donde se añade la macro.

---

## Instalación

1. Abra el archivo de Excel.
2. Presione `Alt + F11` para abrir el Editor de Visual Basic.
3. En el Explorador de Proyectos, localice **Hoja1** y haga doble clic.
4. Copie y pegue el contenido de la macro dentro del módulo de código de **Hoja1**.
5. Guarde el proyecto y cierre el Editor de VBA.

---

## Resumen de Funcionalidades

- **Escritura automática en columna F**: Cuando se ingresa el valor `1013051` en columna E, se escribe `50` automáticamente en la columna F de la misma fila.
- **Protección de valor en columna F**: Si se intenta modificar el `50` en la columna F para celdas donde E = `1013051`, el valor se revierte automáticamente a `50`.
- **Duplicado de valor en B4 → C4**: Cada vez que se cambia la celda B4, su valor se copia en C4.
- **Clasificación de códigos en columna P**: Según el valor ingresado en E (filas 8 en adelante), se asigna etiqueta `V4` o `A4` en la columna P.
- **Formato numérico en columna G**: Todas las celdas con datos en G (desde fila 8) reciben el formato `#,##0.00`.
- **Extracción y mapeo en columna N**: Busca en columna J la cadena "COMPROB GTO CEDIS" seguida de un código, y utiliza matrices predefinidas para mapear a un valor en la columna N.
- **Resaltado condicional en columna I**: Si el texto en columna I (filas 8 en adelante) excede 16 caracteres, se resalta la celda en naranja.

---

## Detalle de Módulos y Subs

### 1. `escribirAutomaticamente`

```vb
Private Sub escribirAutomaticamente(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hoja1")
    
    ' Verifica cambios en columna E
    If Not Intersect(Target, ws.Columns("E")) Is Nothing Then
        If Target.Value = 1013051 Then
            Target.Offset(0, 1).Value = 50
        End If
    End If
End Sub
```
- **Solicitud**: Necesitaban que se hiciera un validador para la cuenta 1013051, ya que era extremadamente importante que unicamente se le pusiera una Posting Key de 50, los demas valores pueden variar con la cuenta, pero solo en 1013051 se exigio cero error, se propuso como solución que una vez que se escriba la cuenta fuera 1013051 cambiara el valor en automatico en Posting Key, y en caso de error humano que lo cambien, iba a seguir reescribiendo el 50
- **Propósito**: Detecta cambios en **columna E**. Si el valor es `1013051`, escribe `50` en la columna F de la misma fila.

### 2. `RevertValueInF`

```vb
Sub RevertValueInF(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hoja1")
    
    ' Si cambia columna F y E = 1013051
    If Not Intersect(Target, ws.Columns("F")) Is Nothing Then
        Dim row As Long: row = Target.Row
        If ws.Cells(row, 5).Value = 1013051 Then
            Application.EnableEvents = False
            Target.Value = 50
            Application.EnableEvents = True
        End If
    End If
End Sub
```
- **Solicitud**: Era para que no pudieran eliminar el 50 previamente agregado por la macro para mitigar el error humano
- **Propósito**: Previene que un usuario modifique el `50` en columna F cuando E = `1013051`, revirtiendo el cambio.

### 3. `Worksheet_Change`
- **Solicitud**: Necesitaban que se diera valores automatizados para los cedis ya que se reconocía el segmento que tienen por defecto, por esto se desarrollo así
Este evento central llama a las dos rutinas anteriores y maneja otras lógicas adicionales:

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    Call escribirAutomaticamente(Target)
    Call RevertValueInF(Target)
    
    ' 1. B4 → C4
    If Not Intersect(Target, Me.Range("B4")) Is Nothing Then
        Application.EnableEvents = False
        Me.Range("C4").Value = Me.Range("B4").Value
        Application.EnableEvents = True
    End If

    ' 2. Clasificación en P (E8:E#)
    If Not Intersect(Target, Me.Range("E8:E1048576")) Is Nothing Then
        Application.EnableEvents = False
        For Each c In Target
            Select Case c.Value
                Case 1013057, 1013058, 1013059, 1013062, 1013512, 1013515, 1013911, 1030401, 1013350
                    Me.Cells(c.Row, "P").Value = "V4"
                Case 2010400, 2010402, 2010403, 2010404, 2010415, 2010606, 2022001
                    Me.Cells(c.Row, "P").Value = "A4"
                Case Else
                    Me.Cells(c.Row, "P").Value = ""
            End Select
        Next c
        Application.EnableEvents = True
    End If

    ' 3. Formato en G
    Range("G8:G" & Cells(Rows.Count, "G").End(xlUp).Row).NumberFormat = "#,##0.00"

    ' 4. Extracción y mapeo en N (columna J)
    valores = Array("301", "302", "304", "305", "307", "308", "309", "310", "311", "313", "317", "319", "331")
    resultados = Array("MXV0005", "MXV0020", "MXV0025", "MXV0019", "MXV0016", "MXV3068", "MXV3056", "MXV0049", "MXV0001", "MXV0017", "MXV0007", "MXV3072", "MXV0004")
    If Not Intersect(Target, Me.Columns("J")) Is Nothing Then
        Application.EnableEvents = False
        lastRow = Cells(Rows.Count, "J").End(xlUp).Row
        count = 0
        For i = 8 To lastRow
            texto = Replace(Cells(i, "J").Value, "'", "")
            If InStr(texto, "COMPROB GTO CEDIS") > 0 Then
                codigo = Mid(texto, InStr(texto, "COMPROB GTO CEDIS") + Len("COMPROB GTO CEDIS") + 1)
                If IsNumeric(codigo) Then
                    For indice = LBound(valores) To UBound(valores)
                        If valores(indice) = codigo Then
                            Cells(i, "N").Value = resultados(indice): Exit For
                        End If
                    Next indice
                Else
                    Cells(i, "N").Value = ""
                End If
            Else
                Cells(i, "N").Value = ""
            End If
        Next i
        Application.EnableEvents = True
    End If

    ' 5. Resaltado en I
    If Not Intersect(Target, Me.Range("I8:I" & Me.Cells(Me.Rows.Count, "I").End(xlUp).Row)) Is Nothing Then
        Application.EnableEvents = False
        For Each celda In Intersect(Target, Me.Range("I8:I" & Me.Cells(Rows.Count, "I").End(xlUp).Row))
            If Len(celda.Value) > 16 Then celda.Interior.Color = RGB(255,165,0) Else celda.Interior.ColorIndex = xlNone
        Next celda
        Application.EnableEvents = True
    End If
End Sub
```  

