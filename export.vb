Sub zobrazeni_vstup_list()
    sheets("nstat12").Visible = True
    sheets("revolvingy").Visible = True
    sheets("nová produkce").Visible = True
    sheets("nstat3").Visible = True
    sheets("Plán_kod_jmeno").Visible = True
    sheets("Produkty_kod_jmeno").Visible = True

End Sub

Sub zamek()
    UserForm1.Show
    
End Sub

Sub A_EXPORT_REGION()
    regionPicker.Show

End Sub





