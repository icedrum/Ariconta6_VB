Attribute VB_Name = "A_versiones"

'------------------------------------------------------
'VERSION 4.0.2
'------------------------------------------------------
'
' Cambios en el 347.
' Tb en el listado de cuentas con datos fiscales
'     ---> Modifacaion en BD usuarios zCuentas
'     ---> Falta tb revisar cta expl consolidada: ztmpCtaexpC
'
'
'------------------------------------------------------





' Cambios en el Listado de facturas
' Aparece un campo msa que es el importe de retencion
'     ---> Modifacaion en BD usuarios zfact prov y cli
'
'------------------------------------------------------




'------------------------------------------------------
'VERSION 4.0.3
'------------------------------------------------------



'Cambios en Venta/Baja inmovilizado



'------------------------------------------------------
'VERSION 4.1...
'------------------------------------------------------
' Junio 2004
'El hecho de guardar los datos del balance como NULL era para k no los pintara.
'El problema son las sumas  por que si hay valores nulos "PETA"
'Lo usaremos para apertura, anteriores y periodo, NO para Saldo
' Para solucionar las dos cosas guardaremos el valor sea cual sea(0 si hace falta) y en la impresion
'desde CRystal ya le diremos k no pinte los 0's

'Antes
'-----
''''''    If ImpD = 0 Then
''''''        d = "NULL"
''''''        Else
''''''        d = TransformaComasPuntos(CStr(ImpD))
''''''    End If
''''''    If ImpH = 0 Then
''''''        H = "NULL"
''''''    Else
''''''        H = TransformaComasPuntos(CStr(ImpH))
''''''    End If

'AHORA
'------
'''''       d = TransformaComasPuntos(CStr(ImpD))

'''''       H = TransformaComasPuntos(CStr(ImpH))


'------------------------------------------------------------------






