Private Function FormatField(FieldName As _
    String, FieldType As ADODB.DataTypeEnum, _
    FieldValue as Variant)
 
    If Len(FieldName) = 0 Then
        Err.Raise vbObjectError + 2007, , _
            "No Field Name was specified for record."
    Else
        Select Case FieldType
            Case ADODB.adChar
                m_ValueString = m_ValueString & _
                    """ & FieldValue & "", "
            Case ADODB.adWChar
                m_ValueString = m_ValueString & _
                    """ & FieldValue & "", "
            Case ADODB.adVarChar
                m_ValueString = m_ValueString & _
                    """ & FieldValue & "", "
            Case ADODB.adVarWChar
                m_ValueString = m_ValueString & _
                    """ & FieldValue & "", "
            Case ADODB.adLongVarWChar
                m_ValueString = m_ValueString & _
                    """ & FieldValue & "", "
            Case ADODB.adDate
                m_ValueString = m_ValueString & _
                    """ & FieldValue & "", "
            Case ADODB.adInteger
                m_ValueString = m_ValueString & _
                    """ FieldValue & ", "
        End Select
    End If
