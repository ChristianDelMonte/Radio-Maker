VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFastCollection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =======================================================
'
' rmUserControlsXP
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' Some graphical UserControls for your VB application.
'
' Copyright � 2006-2007 by Alain Descotes.
'
' rmUserControlsXP is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
'
' rmUserControlsXP is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
'
' =======================================================

Option Explicit


'=======================================================
'APIS
'=======================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


'=======================================================
'variables priv�es
'=======================================================
Private lCount As Long  'nombre d'Items
Private lUbound As Long 'taille du tableau
Private lPtr() As Long  'contient la liste des pointeurs vers les objets

Private Const USTEP As Long = 1000  'variable arbitraire pour le redimensionnement
'du tableau. Plus c'est grand, plus �� gaspille en m�moire (potentiellement)
'mais plus c'est rapide.



'=======================================================
'PUBLIC FUNCTIONS & PROPERTIES
'=======================================================

'=======================================================
'ajoute un item
'=======================================================
Public Sub Add(Obj As Object, Optional ByVal Index As Long = -1)
Attribute Add.VB_Description = "Add an item to collection"
Dim Tmp As Object
    
    'r�cup�re l'objet pass� en param�tre
    Set Tmp = Obj
    
    'on teste si c'est n�cessaire de redimensionner le tableau ou pas
    If lCount = lUbound Then
    
        'alors la taille du tableau est trop petite avec le nouvel Item
        'donc on redimensionne ==> change le lUbound
        lUbound = lUbound + USTEP
        
        'on redimensionne avec Preserve avec la nouvelle taille
        ReDim Preserve lPtr(lUbound)
        
    End If
        
    'incr�mente le nombre d'objets
    lCount = lCount + 1
    
    'si l'index est incorrect ==> alors c'est � la suite
    If Index <= 0 Or Index > lCount Then

        'recup�re le pointeur sur l'objet et le stocke dans le tableau
        lPtr(lCount) = ObjPtr(Obj)

    Else
    
        'alors on ins�re dans la liste � l'Index voulu

        'on copie la partie du tableau APRES index en d�calant d'une case
        'le tableau, gr�ce � Copymemory
        Call CopyMemory(lPtr(Index + 1), lPtr(Index), (lCount - Index) * 4)
    
        'on ins�re le nouvel item au milieu du tableau
        lPtr(Index) = ObjPtr(Obj)
                
    End If
    
    '# On supprime le lien qui existe (ObjPtr) entre Tmp et notre objet.
    '# Ainsi, la r�f�rence que nous venons de placer dans le tableau lPtr restera valide une fois la proc�dure
    '# termin�e (une fois Tmp en dehors de la port�e) : nous avons subtilis�e une r�f�rence vers notre objet.
    Call CopyMemory(Tmp, 0&, 4&)
    
End Sub

'=======================================================
'enl�ve un item
'=======================================================
Public Sub Remove(Index As Long)
Attribute Remove.VB_Description = "Remove an item"
Dim Obj As Object
    
    'si l'index est correct, on supprime
    If Index > 0 And Index <= lCount Then

        '//lib�re l'objet correspondant � l'index � virer
        'copie l'Item dans la variable temporaire gr�ce � son adresse m�moire
        Call CopyMemory(Obj, lPtr(Index), 4)
        'destroy l'Item par Set=Nothing
        Set Obj = Nothing
    
        'on d�cale de 1 vers la gauche le tableau
        Call CopyMemory(lPtr(Index), lPtr(Index + 1), (lCount - Index) * 4)
    
        'change le nombre d'Items
        lCount = lCount - 1
        
        'redimensionne le tableau SI NECESSAIRE
        If lUbound - USTEP > lCount Then
            
            'on enl�ve USTEP items car les USTEP derniers ne servent plus
            lUbound = lUbound - USTEP
            
            'alors on redimensionne le tableau en virant USTEP �l�ments
            ReDim Preserve lPtr(lUbound)
        
        End If
        
    End If
    
End Sub

'=======================================================
'clear la collection
'=======================================================
Public Sub Clear()
Attribute Clear.VB_Description = "Clear the collection"
Dim ptrFirst As Long
Dim x As Long
Dim Obj As Object
    
    'que si des Items pr�sents
    If lCount Then
        
        'supprime tous les objets de la liste.
        'ne parcours pas toute la liste : ne r�cup�re que le pointeur du
        'premier item de la liste et destroy les items situ�s � chaque pointeur
        'Chaque �l�ment du tableau est d�cal� de 4 bytes (Long) en m�moire
        
        'r�cup�re le pointeur du premier objet
        ptrFirst = VarPtr(lPtr(1))
        
        For x = ptrFirst To ptrFirst + (lCount - 1) * 4 Step 4
            
            'on r�cup�re l'objet dans la variable temp Obj
            Call CopyMemory(Obj, ByVal x, 4&) '4 <==> long
            
            'destroy l'item par Set=Nothing
            Set Obj = Nothing
            
        Next x
        
        'supprime le tableau et initialise les valeurs
        ReDim lPtr(0)
        lCount = 0: lUbound = 0
    
    End If
 
End Sub

'=======================================================
'r�cup�re l'item point� par l'index
'=======================================================
Public Function Item(Index As Long) As Object
Attribute Item.VB_Description = "Get an item"
Attribute Item.VB_UserMemId = 0
Dim Obj As Object
    
    'si l'index est valide
    If Index > 0 And Index <= lCount Then
        
        'on a besoin de r�cup�rer notre objet dans une variable temporaire
        Call CopyMemory(Obj, lPtr(Index), 4)      '4 octets
        
        'r�cup�re l'objet en r�sultat de la function
        Set Item = Obj
        
        'd�truit l'objet temporaire maintenant inutile
        Call CopyMemory(Obj, 0&, 4)
        
    End If

End Function

'=======================================================
'r�cup�re le Count de la collection
'=======================================================
Public Property Get Count() As Long: Count = lCount: End Property
Attribute Count.VB_Description = "Get the number of items in collection"





'=======================================================
'SUB DE LA CLASSE
'=======================================================

'=======================================================
'initialisation et destruction
'=======================================================
Private Sub Class_Initialize()
    lCount = 0: lUbound = 0: ReDim lPtr(0)
End Sub
Private Sub Class_Terminate()
    Call Clear
    Erase lPtr  'supprime compl�tement le tableau
End Sub
