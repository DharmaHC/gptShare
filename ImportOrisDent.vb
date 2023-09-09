Imports System.Text

Module ImportOrisDent
    Private dt As DataTable
    Public ErrorOccurred As Boolean
    Private MyOwnerForm As FrmInvoicesStatements
    Private allWorkareasList As List(Of Workarea)
    Private ImportPath As String = "\\172.16.17.30\mdata\Export"

    Public Sub DoImport(refWorkareasList As List(Of Workarea), sender As FrmInvoicesStatements)
        MyOwnerForm = sender
        MyOwnerForm.Cursor = Cursors.WaitCursor
        ErrorOccurred = False
        allWorkareasList = refWorkareasList
        ImportTestataData()
        If ErrorOccurred Then Exit Sub
        ImportFattureEmesseData()
        If ErrorOccurred Then Exit Sub
        ImportPrimaNota()
        MyOwnerForm.Cursor = Cursors.Default
    End Sub


    Private Sub ImportTestataData()


        Dim fileName As String = ImportPath & "\ASTER_TST.txt"

        Dim LastName As String = String.Empty
        Dim FirstName As String = String.Empty
        Dim pdfString As String = String.Empty

        dt = New DataTable
        ' LOGIC for Retrieving Data Table...

        dt.Columns.Add("Num_Doc")
        dt.Columns.Add("Anno_Doc")
        dt.Columns.Add("Codice_Doc")
        dt.Columns.Add("Data_Doc")
        dt.Columns.Add("Totale_Doc")
        dt.Columns.Add("TotaleNetto_Doc")
        dt.Columns.Add("Bollo")
        dt.Columns.Add("IVA")
        dt.Columns.Add("Franchigia")
        dt.Columns.Add("Totale_Pagato")
        dt.Columns.Add("Totale_da_Pagare")
        dt.Columns.Add("Id_Accettazione")
        dt.Columns.Add("RagioneSoc_o_Intestatario")
        dt.Columns.Add("Indirizzo")
        dt.Columns.Add("Citta")
        dt.Columns.Add("Cod_Fiscale")
        dt.Columns.Add("PartitaIVA")
        dt.Columns.Add("Cap")
        dt.Columns.Add("Flg_Nota_di_Credito")
        dt.Columns.Add("Flg_HA_Nota_di_Credito")
        dt.Columns.Add("Flg_HA_Bollo")
        dt.Columns.Add("Flg_Intestata_a_Paziente")
        dt.Columns.Add("Flg_Intestata_ad_Ente")
        dt.Columns.Add("Num_Fattura_Madre")
        dt.Columns.Add("Anno_Fattura_Madre")
        dt.Columns.Add("DocCode_Fattura_Madre")
        dt.Columns.Add("ApplicantId")
        dt.Columns.Add("Flg_HA_Iva")
        dt.Columns.Add("Iva_Percent")
        dt.Columns.Add("Utente")
        dt.Columns.Add("Id_Tipo_Pagamento")
        dt.Columns.Add("MaskedCode")
        dt.Columns.Add("Tipo_Documento")
        dt.Columns.Add("Cognome_Paz")
        dt.Columns.Add("Nome_Paz")

        Dim rowCounter As Integer = 0
        Using reader As System.IO.StreamReader = New System.IO.StreamReader(fileName)
            While Not reader.EndOfStream
                Try
                    Dim row As String = reader.ReadLine
                    rowCounter += 1

                    Dim dr As DataRow = dt.NewRow
                    dr("Num_Doc") = CInt((Mid(row, 1, 10)))
                    dr("Anno_Doc") = Mid(row, 11, 4)
                    dr("Codice_Doc") = Trim(Mid(row, 15, 10))
                    If CStr(dr("Codice_Doc")).Trim = "M" Then dr("Codice_Doc") = "O"
                    dr("Data_Doc") = Mid(row, 25, 8)
                    dr("Totale_Doc") = Mid(row, 33, 10)
                    dr("TotaleNetto_Doc") = Mid(row, 43, 10)
                    dr("Bollo") = Mid(row, 53, 5)
                    If CDbl(dr("Bollo")) = 0 Then
                        dr("Bollo") = Nothing
                    End If
                    dr("IVA") = Mid(row, 58, 10)
                    If CDbl(dr("IVA")) = 0 Then
                        dr("IVA") = Nothing
                    End If
                    dr("Franchigia") = Mid(row, 68, 10)
                    If CStr(dr("Franchigia")).Trim = "" Or CDbl(dr("Franchigia")) = 0 Then dr("Franchigia") = Nothing
                    dr("Totale_da_Pagare") = Mid(row, 78, 10)
                    dr("Totale_Pagato") = Mid(row, 88, 10)
                    dr("Id_Accettazione") = Mid(row, 98, 8)
                    dr("RagioneSoc_o_Intestatario") = Mid(row, 106, 100)
                    dr("Indirizzo") = Mid(row, 206, 100)
                    dr("Citta") = Mid(row, 306, 50)
                    dr("Cod_Fiscale") = Mid(row, 356, 16)
                    dr("PartitaIVA") = Mid(row, 372, 16)
                    dr("Cap") = Mid(row, 388, 5)
                    dr("Flg_Nota_di_Credito") = Mid(row, 393, 1)
                    dr("Flg_HA_Nota_di_Credito") = Mid(row, 394, 1)
                    dr("Flg_HA_Bollo") = Mid(row, 395, 1)
                    dr("Flg_Intestata_a_Paziente") = Mid(row, 396, 1)
                    dr("Flg_Intestata_ad_Ente") = Mid(row, 397, 1)
                    dr("Num_Fattura_Madre") = Mid(row, 398, 10)
                    If CStr(dr("Num_Fattura_Madre")).Trim = "" Then dr("Num_Fattura_Madre") = Nothing
                    dr("Anno_Fattura_Madre") = Mid(row, 408, 4)
                    If CStr(dr("Anno_Fattura_Madre")).Trim = "" Then dr("Anno_Fattura_Madre") = Nothing
                    dr("DocCode_Fattura_Madre") = Mid(row, 412, 10)
                    If CStr(dr("DocCode_Fattura_Madre")).Trim = "M" Then
                        dr("DocCode_Fattura_Madre") = "O"
                    End If
                    If CStr(dr("DocCode_Fattura_Madre")).Trim = "" Then dr("DocCode_Fattura_Madre") = Nothing
                    dr("ApplicantId") = Mid(row, 422, 10)
                    dr("Flg_HA_Iva") = Mid(row, 432, 1)
                    If dr("Flg_HA_Iva") = 1 Then
                        Dim popo = 1
                    End If
                    dr("Iva_Percent") = Mid(row, 433, 2)
                    dr("Utente") = Mid(row, 435, 50)
                    dr("Id_Tipo_Pagamento") = Mid(row, 485, 50)
                    dr("MaskedCode") = Mid(row, 535, 50)
                    If InStr(CStr(dr("MaskedCode")), "/M") Then
                        dr("MaskedCode") = CStr(dr("MaskedCode")).Replace("/M", "/O")
                    End If
                    dr("Tipo_Documento") = Mid(row, 585, 2)
                    dr("Cognome_Paz") = Mid(row, 587, 100)
                    dr("Nome_Paz") = Mid(row, 687, 100)

                    ' Get receiver name
                    LastName = dr("Cognome_Paz")
                    FirstName = dr("Nome_Paz")

                    ' Get Pdf and add to table
                    pdfString = Mid(row, 787, row.Length - 786)
                    If pdfString IsNot Nothing Then
                        Using tmp_dbContext As New HealthNET_DataEntities
                            Dim newInvoicePdf As New InvoicesPdf With {
                                .InvoiceNumber = dr("Num_Doc"),
                                .InvoiceYear = dr("Anno_Doc"),
                                .DocCode = dr("Codice_Doc"),
                                .PdfString = pdfString}

                            Dim ExistingPdf As InvoicesPdf = tmp_dbContext.InvoicesPdfs.Where(
                                Function(ip) ip.InvoiceNumber = newInvoicePdf.InvoiceNumber And
                                ip.InvoiceYear = newInvoicePdf.InvoiceYear And
                                ip.DocCode.Trim = newInvoicePdf.DocCode.Trim).SingleOrDefault

                            If ExistingPdf IsNot Nothing Then
                                tmp_dbContext.InvoicesPdfs.Remove(ExistingPdf)
                            End If
                            tmp_dbContext.InvoicesPdfs.Add(newInvoicePdf)
                            tmp_dbContext.SaveChanges()
                        End Using
                    End If
                    dt.Rows.Add(dr)
                Catch ex As Exception
                    MyOwnerForm.Cursor = Cursors.Default
                    WaitingDialog.CloseDialog()
                    Dim ErrorBody As String = "Errore nell'acquisizione del file alla riga " & rowCounter & vbCrLf &
                                ex.ToString
                    MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                    ErrorOccurred = True
                    Exit Sub
                End Try
            End While
        End Using

        Dim dbContext As New HealthNET_DataEntities
        Dim counter As Integer = 0
        For Each r In dt.Rows
            counter += 1
            WaitingDialog.ShowDialog("Importo e creo fattura " & counter & " di " & dt.Rows.Count)
            Try
                Dim newInvoice As Invoice
                Dim isUpdate As Boolean = False

                Dim invoiceNumber As Integer = r("Num_Doc")
                Dim invoiceYear As Integer = r("Anno_Doc")
                Dim invoiceDocCode As String = r("Codice_Doc")
                If invoiceDocCode.Trim = "M" Then invoiceDocCode = "O"
                If dbContext.Invoices.Where(Function(inv) inv.InvoiceNumber = invoiceNumber And
                                            inv.InvoiceYear = invoiceYear And inv.DocCode.Trim = invoiceDocCode.Trim).Any Then

                    isUpdate = True
                    newInvoice = dbContext.Invoices.Where(Function(inv) inv.InvoiceNumber = invoiceNumber And
                                            inv.InvoiceYear = invoiceYear And inv.DocCode.Trim = invoiceDocCode.Trim).Single
                Else
                    newInvoice = New Invoice
                End If

                newInvoice.InvoiceNumber = r("Num_Doc")
                newInvoice.InvoiceYear = r("Anno_Doc")
                If isUpdate = False Then
                    newInvoice.DocCode = r("Codice_Doc")
                End If
                newInvoice.InvoiceDate = New Date(Mid(r("Data_Doc"), 5, 4), Mid(r("Data_Doc"), 3, 2), Mid(r("Data_Doc"), 1, 2))
                newInvoice.Amount = CDec(r("Totale_Doc"))
                newInvoice.NetAmount = CDec(r("TotaleNetto_Doc"))
                If r("Bollo") IsNot Nothing AndAlso r("Bollo") IsNot DBNull.Value Then
                    newInvoice.AdditionalTax = CDec(r("Bollo"))
                End If
                If r("IVA") IsNot Nothing AndAlso r("IVA") IsNot DBNull.Value Then
                    newInvoice.VAT = CDec(r("IVA"))
                End If
                If newInvoice.VAT = 0 Then newInvoice.VAT = Nothing
                If r("Iva_Percent") IsNot Nothing AndAlso r("Iva_Percent") IsNot DBNull.Value Then
                    newInvoice.VAT_Percentage = CDec(r("Iva_Percent"))
                End If
                If r("Franchigia") IsNot Nothing AndAlso r("Franchigia") IsNot DBNull.Value Then
                    newInvoice.AgreementDeductible = CDec(r("Franchigia"))
                End If
                newInvoice.AmountPaid = CDec(r("Totale_Pagato"))
                newInvoice.AmountToPay = CDec(r("Totale_da_Pagare"))

                newInvoice.ReceiverName = r("RagioneSoc_o_Intestatario").ToString.Trim
                newInvoice.ReceiverAddress = r("Indirizzo").ToString.Trim
                newInvoice.ReceiverCity = r("Citta").ToString.Trim
                newInvoice.ReceiverTaxCode = r("Cod_Fiscale")
                newInvoice.ReceiverCompanyTaxCode = r("PartitaIVA")
                newInvoice.ReceiverZipCode = r("Cap")
                newInvoice.IsCreditNote = If(r("Flg_Nota_di_Credito") = "1", True, False)
                newInvoice.HasCreditNote = If(r("Flg_HA_Nota_di_Credito") = "1", True, False)
                newInvoice.FixedAdditionalAmountOnInvoice = If(r("Flg_HA_Bollo") = "1", True, False)
                newInvoice.IsPatientInvoiceRecipient = If(r("Flg_Intestata_a_Paziente") = "1", True, False)
                newInvoice.IsApplicantInvoiceRecipient = If(r("Flg_Intestata_ad_Ente") = "1", True, False)
                newInvoice.CreditNoteParentInvoice = If(r("Num_Fattura_Madre") IsNot Nothing AndAlso r("Num_Fattura_Madre").ToString.Trim <> "", CInt(r("Num_Fattura_Madre")), Nothing)
                If newInvoice.CreditNoteParentInvoice = 0 Then
                    newInvoice.CreditNoteParentInvoice = Nothing
                End If
                newInvoice.CreditNoteParentInvoiceYear = If(r("Anno_Fattura_Madre") IsNot Nothing AndAlso r("Anno_Fattura_Madre").ToString.Trim <> "", CInt(r("Anno_Fattura_Madre")), Nothing)
                If newInvoice.CreditNoteParentInvoiceYear = 0 Then
                    newInvoice.CreditNoteParentInvoiceYear = Nothing
                End If
                newInvoice.CreditNoteParentInvoiceDocCode = If(r("DocCode_Fattura_Madre") IsNot Nothing AndAlso r("DocCode_Fattura_Madre").ToString.Trim <> "", r("DocCode_Fattura_Madre"), Nothing)
                newInvoice.ApplicantId = r("ApplicantId").ToString.Trim
                If newInvoice.ApplicantId IsNot Nothing Then
                    Try
                        'newInvoice.Applicant = dbContext.Applicants.Where(Function(a) a.ApplicantId.Trim = newInvoice.ApplicantId.Trim).Single
                    Catch ex As Exception

                    End Try
                End If
                newInvoice.InvoiceHasVAT = If(r("Flg_HA_Iva") = "1", True, False)
                newInvoice.VAT_Percentage = If(r("Iva_Percent") IsNot Nothing, CInt(r("Iva_Percent")), Nothing)
                newInvoice.IssuedBy = r("Utente").ToString.Trim
                If newInvoice.InvoiceNumber = 1448 Then
                    Dim popo = 2
                End If
                newInvoice.CashFlowPaymentModeId = r("Id_Tipo_Pagamento").ToString.Trim
                newInvoice.MaskedCounter = r("MaskedCode").ToString.Trim
                newInvoice.DocTypeId = r("Tipo_Documento").ToString.Trim
                newInvoice.CompanyId = "ASTER"

                If newInvoice.IsPatientInvoiceRecipient Then
                    Dim codFiscalePaz As String = newInvoice.ReceiverTaxCode.Trim
                    Dim patDbContext As New HealthNET_DataEntities
                    Dim Pat As Patient = Nothing
                    Dim patientExists As Boolean = CheckPatientExistance(codFiscalePaz, Pat, patDbContext)
                    If Not patientExists Then
                        Pat = New Patient
                        With Pat
                            Try
                                ReverseTaxCode_Computing(codFiscalePaz, Pat, patDbContext)
                            Catch ex As Exception
                                MyOwnerForm.Cursor = Cursors.Default
                                WaitingDialog.CloseDialog()
                                Dim ErrorBody As String = "Errore nella verifica del codice fiscale " & codFiscalePaz & vbCrLf &
                                    "Fattura " & newInvoice.InvoiceNumber & vbCrLf & ex.ToString
                                MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                                ErrorOccurred = True
                                Exit Sub
                            End Try
                            .Address = newInvoice.ReceiverAddress
                            .City = newInvoice.ReceiverCity
                            Try
                                .CityCode = patDbContext.Cities.Where(Function(c) c.City1.Trim = .City.Trim).First.CodIstatCity
                            Catch ex As Exception

                            End Try
                            .CodiceFiscale = codFiscalePaz

                            ' Try to get receiver name
                            .LastName = r("Cognome_Paz").ToString.Trim
                            .FirstName = r("Nome_Paz").ToString.Trim

                            .IsFakePatient = False
                            .PatientReferentId = "MET"
                            .ZipCode = newInvoice.ReceiverZipCode
                        End With
                        patDbContext.Patients.Add(Pat)
                        patDbContext.SaveChanges()
                    Else
                        With Pat
                            Try
                                ReverseTaxCode_Computing(codFiscalePaz, Pat, patDbContext)
                            Catch ex As Exception
                                MessageBox.Show("Verificare il codice fiscale del paziente: " & .LastName & " " & .FirstName & vbCrLf &
                                                "Importazione effettuata ma verifica del codice fiscale da riscontrare", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            End Try
                            .Address = newInvoice.ReceiverAddress
                            .City = newInvoice.ReceiverCity
                            Try
                                .CityCode = patDbContext.Cities.Where(Function(c) c.City1.Trim = .City.Trim).First.CodIstatCity
                            Catch ex As Exception
                                MessageBox.Show("Errore nel campo città del paziente: " & .LastName & " " & .FirstName & vbCrLf & vbCrLf &
                                                "Fattura n. " & newInvoice.InvoiceNumber & ". Verificare il dato", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                Continue For
                            End Try
                            .CodiceFiscale = codFiscalePaz

                            ' Try to get receiver name
                            .LastName = r("Cognome_Paz").ToString.Trim
                            .FirstName = r("Nome_Paz").ToString.Trim
                            .CreatedBy = "Methodent"
                            .CreationDate = Now
                            .IsFakePatient = False
                            .PatientReferentId = "MET"
                            .ZipCode = newInvoice.ReceiverZipCode

                            ' Check Primary doctor existance, if not set PrimaryCareDoctor = nothing
                            If Pat.PrimaryCareDoctorCode IsNot Nothing Then
                                Dim ExistingPrimaryCareDoctor As PrimaryCareDoctor = patDbContext.PrimaryCareDoctors.Where(Function(pcd) pcd.PrimaryCareDoctorId.Trim = Pat.PrimaryCareDoctorCode.Trim).FirstOrDefault
                                If ExistingPrimaryCareDoctor Is Nothing Then
                                    Pat.PrimaryCareDoctorCode = Nothing
                                End If
                            End If
                        End With
                        'patDbContext.Patients.Add(Pat)
                        patDbContext.SaveChanges()
                    End If

                    Dim RegistrationDate As DateTime = newInvoice.InvoiceDate
                    RegistrationDate = RegistrationDate.AddHours(Now.Hour)
                    RegistrationDate = RegistrationDate.AddMinutes(Now.Minute)
                    RegistrationDate = RegistrationDate.AddSeconds(Now.Second)

                    Dim RegistrationId As Integer = 0
                    If newInvoice.ExaminationId Is Nothing Or Pat.ExaminationsAndConsultations.Count = 0 Or
                        patDbContext.ExaminationsAndConsultations.Where(Function(eac) eac.PatientId = Pat.PatientId And
                        eac.ExaminationId = newInvoice.ExaminationId).Any = False Then
                        ' Create new ExaminationsAndConsultation
                        CreateNewRegistration_ForGivenPatient(Pat, patDbContext, Nothing, Nothing, RegistrationDate, "MET", RegistrationId)
                    Else
                        RegistrationId = newInvoice.ExaminationId
                    End If

                    Dim currentExamination As ExaminationsAndConsultation = Nothing
                    If RegistrationId = 0 Then
                        WaitingDialog.CloseDialog()
                        MyOwnerForm.Cursor = Cursors.Default
                        Dim ErrorBody As String = "Errore nella creazione dell'accettazione. Documento " & r("MaskedCode") & vbCrLf &
                                "Paziente: " & Pat.LastName & "" & Pat.FirstName
                        MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                        ErrorOccurred = True
                        Exit Sub
                    Else
                        newInvoice.ExaminationId = RegistrationId
                        Try
                            currentExamination = Pat.ExaminationsAndConsultations.Where(Function(eac) eac.ExaminationId = RegistrationId).Single
                        Catch ex As Exception

                        End Try
                    End If

                    If currentExamination Is Nothing Then
                        WaitingDialog.CloseDialog()
                        MyOwnerForm.Cursor = Cursors.Default
                        Dim ErrorBody As String = "Errore nella definizione dell'accettazione. Documento " & r("MaskedCode") & vbCrLf &
                                "Paziente: " & Pat.LastName & " " & Pat.FirstName
                        MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                        ErrorOccurred = True
                        Exit Sub
                    End If


                    patDbContext.Database.ExecuteSqlCommand("UPDATE ExaminationsAndConsultations SET InvoicedStatus = 4 WHERE ExaminationId = " & currentExamination.ExaminationId)

                    newInvoice.ExaminationId = currentExamination.ExaminationId
                    patDbContext.SaveChanges()

                End If

                If newInvoice.IsCreditNote Then
                    newInvoice.Amount = -newInvoice.Amount
                    newInvoice.NetAmount = -newInvoice.NetAmount
                    newInvoice.AdditionalTax = -newInvoice.AdditionalTax
                    newInvoice.AmountPaid = -newInvoice.AmountPaid
                    newInvoice.AmountToPay = -newInvoice.AmountToPay
                End If

                Dim InvoicesWithDetailsToImport As New List(Of InvoicesWithDetails)
                Dim newIwd As New InvoicesWithDetails
                newIwd.AdditionalTax = newInvoice.AdditionalTax
                newIwd.AgreementDeductible = newInvoice.AgreementDeductible
                newIwd.Amount = newInvoice.Amount
                newIwd.AmountPaid = newInvoice.AmountPaid
                newIwd.AmountToPay = newInvoice.AmountToPay
                newIwd.Applicant = newInvoice.Applicant
                newIwd.ApplicantId = newInvoice.ApplicantId
                newIwd.CashFlowPaymentModeId = newInvoice.CashFlowPaymentModeId
                newIwd.CashFlows = newInvoice.CashFlows
                newIwd.CausalDefault = newInvoice.CausalDefault
                newIwd.CompanyId = newInvoice.CompanyId
                newIwd.CreditNoteParentInvoice = newInvoice.CreditNoteParentInvoice
                newIwd.CreditNoteParentInvoiceDocCode = newInvoice.CreditNoteParentInvoiceDocCode
                newIwd.CreditNoteParentInvoiceYear = newInvoice.CreditNoteParentInvoiceYear
                newIwd.DocCode = newInvoice.DocCode
                newIwd.DocTypeId = newInvoice.DocTypeId
                newIwd.DocumentsSent = newInvoice.DocumentsSent
                newIwd.ExaminationId = newInvoice.ExaminationId
                newIwd.ExportToTS730_Suspended = newInvoice.ExportToTS730_Suspended
                newIwd.FixedAdditionalAmountOnInvoice = newInvoice.FixedAdditionalAmountOnInvoice
                newIwd.HasCreditNote = newInvoice.HasCreditNote
                newIwd.InvoiceAdditionalAmounts = newInvoice.InvoiceAdditionalAmounts
                newIwd.InvoiceDate = newInvoice.InvoiceDate
                newIwd.InvoiceDetails = newInvoice.InvoicesDetails
                newIwd.InvoiceDetailsUnmerged = newInvoice.InvoicesDetailsUnmergeds
                newIwd.InvoiceHasVAT = newInvoice.InvoiceHasVAT
                newIwd.InvoiceNumber = newInvoice.InvoiceNumber
                newIwd.InvoicesDetails = newInvoice.InvoicesDetails
                newIwd.InvoicesDetailsUnmergeds = newInvoice.InvoicesDetailsUnmergeds
                newIwd.InvoicesDocType = newInvoice.InvoicesDocType
                newIwd.InvoiceYear = newInvoice.InvoiceYear
                newIwd.IsApplicantInvoiceRecipient = newInvoice.IsApplicantInvoiceRecipient
                newIwd.IsCancelled = newInvoice.IsCancelled
                newIwd.IsCreditNote = newInvoice.IsCreditNote
                newIwd.IsDeferredInvoice = newInvoice.IsDeferredInvoice
                newIwd.IsPatientInvoiceRecipient = newInvoice.IsPatientInvoiceRecipient
                newIwd.NetAmount = newInvoice.NetAmount
                newIwd.OtherReceiverAddress = newInvoice.OtherReceiverAddress
                newIwd.OtherReceiverAddressNumber = newInvoice.OtherReceiverAddressNumber
                newIwd.OtherReceiverCity = newInvoice.OtherReceiverCity
                newIwd.OtherReceiverName = newInvoice.OtherReceiverName
                newIwd.OtherReceiverPatientId = newInvoice.OtherReceiverPatientId
                newIwd.OtherReceiverTaxCode = newInvoice.OtherReceiverTaxCode
                newIwd.OtherReceiverZipCode = newInvoice.OtherReceiverZipCode
                newIwd.ReceiverAddress = newInvoice.ReceiverAddress
                newIwd.ReceiverAddressNumber = newInvoice.ReceiverAddressNumber
                newIwd.ReceiverCity = newInvoice.ReceiverCity
                newIwd.ReceiverCompanyTaxCode = newInvoice.ReceiverCompanyTaxCode
                newIwd.ReceiverName = newInvoice.ReceiverName
                newIwd.ReceiverTaxCode = newInvoice.ReceiverTaxCode
                newIwd.ReceiverZipCode = newInvoice.ReceiverZipCode
                newIwd.VAT = newInvoice.VAT
                newIwd.VAT_Percentage = newInvoice.VAT_Percentage

                InvoicesWithDetailsToImport.Add(newIwd)
                Dim check As Boolean = FrmInvoicesStatements.CheckInvoices_DataConsistency(InvoicesWithDetailsToImport, "Import Meth")

                If check = False Then Continue For
                If isUpdate = False Then
                    dbContext.Invoices.Add(newInvoice)
                End If
                Try
                    dbContext.SaveChanges()
                Catch ex As Exception
                    MyOwnerForm.Cursor = Cursors.Default
                    WaitingDialog.CloseDialog()
                    Dim ErrorBody As String = "Errore nel salvataggio su db" & vbCrLf &
                                ex.ToString
                    MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                    ErrorOccurred = True
                    Exit Sub
                End Try


            Catch ex As Exception
                WaitingDialog.CloseDialog()
                MyOwnerForm.Cursor = Cursors.Default
                Dim ErrorBody As String = "Errore nell'inserimento in db del documento" & r("MaskedCode") & vbCrLf &
                                ex.ToString
                MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                ErrorOccurred = True
                Exit Sub
            End Try
            WaitingDialog.CloseDialog()


        Next
        WaitingDialog.CloseDialog()

        Try
            dbContext.SaveChanges()
        Catch ex As Exception
            MyOwnerForm.Cursor = Cursors.Default
            WaitingDialog.CloseDialog()
            Dim ErrorBody As String = "Errore nel salvataggio su db " & vbCrLf &
                ex.ToString
            MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            GlobalFunctions.WriteImportFromHostLog(ErrorBody)
            ErrorOccurred = True
            Exit Sub
        End Try
    End Sub

    Private Function CheckPatientExistance(ByVal codFiscalePaz As String, ByRef pat As Patient, ByRef patDbContext As HealthNET_DataEntities) As Boolean
        If patDbContext.Patients.Where(Function(p) p.CodiceFiscale.Trim = codFiscalePaz.Trim).Any Then
            pat = patDbContext.Patients.Where(Function(p) p.CodiceFiscale.Trim = codFiscalePaz.Trim).First
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub ImportFattureEmesseData()
        Dim fileName As String = ImportPath & "\ASTER_DET.txt"

        Dim fileContents As String
        fileContents = My.Computer.FileSystem.ReadAllText(fileName)

        dt = New DataTable
        ' LOGIC for Retrieving Data Table...

        dt.Columns.Add("Num_Doc")
        dt.Columns.Add("Anno_Doc")
        dt.Columns.Add("Codice_Doc")
        dt.Columns.Add("ExamId")
        dt.Columns.Add("ExamName")
        dt.Columns.Add("ExamPrice")
        dt.Columns.Add("Quantity")
        Dim rowCounter As Integer = 0
        Do Until fileContents.Length = 0
            Try
                rowCounter += 1
                Dim row As String = ""

                If InStr(fileContents, vbCrLf) Then
                    row = Strings.Left(fileContents, InStr(fileContents, vbCrLf) - 1)
                    fileContents = Strings.Right(fileContents, Len(fileContents) - InStr(fileContents, vbCrLf) - 1)
                Else
                    row = fileContents
                    fileContents = ""
                End If

                Dim dr As DataRow = dt.NewRow
                dr("Num_Doc") = CInt((Mid(row, 1, 10)))
                dr("Anno_Doc") = Mid(row, 11, 4)
                dr("Codice_Doc") = Trim(Mid(row, 15, 10))
                dr("ExamId") = Mid(row, 25, 10)
                dr("ExamName") = Mid(row, 35, 100)
                dr("ExamPrice") = Mid(row, 135, 10)
                dr("Quantity") = CInt(Mid(row, 145, 3))
                dt.Rows.Add(dr)

            Catch ex As Exception
                MyOwnerForm.Cursor = Cursors.Default
                WaitingDialog.CloseDialog()
                Dim ErrorBody As String = "Errore nell'acquisizione del file dettagli fatture alla riga " & rowCounter & vbCrLf &
                                ex.ToString
                MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                ErrorOccurred = True
                Exit Sub
            End Try
        Loop

        Dim dbContext As New HealthNET_DataEntities

        Dim counter As Integer = 0
        WaitingDialog.ShowDialog("Importo dettagli fatture ")
        For Each r In dt.Rows
            counter += 1

            Dim newInvoiceDetail As InvoicesDetail
            Dim invoiceNumber As Integer = CInt(r("Num_Doc"))
            Dim invoiceYear As Integer = CInt(r("Anno_Doc"))
            Dim invoiceDocCode As String = r("Codice_Doc")
            If invoiceDocCode.Trim = "M" Then invoiceDocCode = "O"

            Dim invoiceExamId As Integer = CInt(r("ExamId"))
            Try
                Dim isUpdate As Boolean = False


                If dbContext.InvoicesDetails.Where(Function(inv) inv.InvoiceNumber = invoiceNumber And
                                            inv.InvoiceYear = invoiceYear And inv.DocCode.Trim = invoiceDocCode.Trim And
                                            inv.ExamId = invoiceExamId).Any Then

                    isUpdate = True
                    newInvoiceDetail = dbContext.InvoicesDetails.Where(Function(inv) inv.InvoiceNumber = invoiceNumber And
                                            inv.InvoiceYear = invoiceYear And inv.DocCode.Trim = invoiceDocCode.Trim And
                                            inv.ExamId = invoiceExamId).Single
                Else
                    newInvoiceDetail = New InvoicesDetail
                End If

                newInvoiceDetail.InvoiceNumber = invoiceNumber
                newInvoiceDetail.InvoiceYear = invoiceYear
                newInvoiceDetail.ExamId = invoiceExamId
                newInvoiceDetail.ExamName = r("ExamName")
                newInvoiceDetail.ExamPrice = CDec(r("ExamPrice"))
                newInvoiceDetail.ExamVersion = 0
                newInvoiceDetail.Quantity = CInt(r("Quantity"))
                'newInvoiceDetail.VAT = r("IVA")
                newInvoiceDetail.RevenueCenterId = "ODO"
                Dim relatedInvoice As Invoice = dbContext.Invoices.Where(Function(inv) inv.InvoiceNumber = invoiceNumber And
                                                                         inv.InvoiceYear = invoiceYear And
                                                                         inv.DocCode.Trim = invoiceDocCode.Trim).SingleOrDefault

                If relatedInvoice IsNot Nothing AndAlso relatedInvoice.IsCreditNote Then
                    newInvoiceDetail.ExamPrice = -newInvoiceDetail.ExamPrice
                End If

                If isUpdate = False Then
                    newInvoiceDetail.DocCode = invoiceDocCode
                    dbContext.InvoicesDetails.Add(newInvoiceDetail)
                End If
            Catch ex As Exception
                WaitingDialog.CloseDialog()
                MyOwnerForm.Cursor = Cursors.Default
                Dim ErrorBody As String = "Errore nell'inserimento in db del dettaglio documento" & invoiceNumber & "/" & invoiceYear & "/" & invoiceDocCode & "/" & vbCrLf &
                                ex.ToString
                MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                ErrorOccurred = True
                Exit Sub
            End Try
            Try
                dbContext.SaveChanges()
            Catch ex As Exception
                WaitingDialog.CloseDialog()
                MyOwnerForm.Cursor = Cursors.Default
                Dim ErrorBody As String = "Errore nel salvataggio su db in Import Fatture emesse" & vbCrLf &
                ex.ToString
                MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                ErrorOccurred = True
                Exit Sub
            End Try
        Next
        WaitingDialog.CloseDialog()

        Try
            dbContext.SaveChanges()
        Catch ex As Exception
            WaitingDialog.CloseDialog()
            MyOwnerForm.Cursor = Cursors.Default
            Dim ErrorBody As String = "Errore nel salvataggio su db in Import Fatture emesse" & vbCrLf &
                ex.ToString
            MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            GlobalFunctions.WriteImportFromHostLog(ErrorBody)
            ErrorOccurred = True
            Exit Sub
        End Try

    End Sub

    Function csvBytesWriter(ByRef dTable As DataTable) As Byte()

        '--------Columns Name---------------------------------------------------------------------------

        Dim sb As StringBuilder = New StringBuilder()
        Dim intClmn As Integer = dTable.Columns.Count

        Dim i As Integer = 0
        For i = 0 To intClmn - 1 Step i + 1
            sb.Append("""" + dTable.Columns(i).ColumnName.ToString() + """")
            If i = intClmn - 1 Then
                sb.Append("")
            Else
                sb.Append(";")
            End If
        Next
        sb.Append(vbNewLine)

        '--------Data By  Columns---------------------------------------------------------------------------

        Dim row As DataRow
        For Each row In dTable.Rows

            Dim ir As Integer = 0
            For ir = 0 To intClmn - 1 Step ir + 1
                If IsNumeric(row(ir)) Then
                    sb.Append(row(ir))
                    If ir = intClmn - 1 Then
                        sb.Append("")
                    Else
                        sb.Append(";")
                    End If
                Else
                    sb.Append("""" + row(ir).ToString().Replace("""", """""") + """")
                    If ir = intClmn - 1 Then
                        sb.Append("")
                    Else
                        sb.Append(";")
                    End If
                End If


            Next
            sb.Append(vbNewLine)
        Next

        Return System.Text.Encoding.UTF8.GetBytes(sb.ToString)

    End Function

    Private Sub ImportPrimaNota()

        Dim fileName As String = ImportPath & "\ASTER_MOV.txt"

        Dim fileContents As String
        fileContents = My.Computer.FileSystem.ReadAllText(fileName)

        dt = New DataTable
        ' LOGIC for Retrieving Data Table...

        dt.Columns.Add("Num_Doc")
        dt.Columns.Add("Anno_Doc")
        dt.Columns.Add("Codice_Doc")
        dt.Columns.Add("Totale_Doc")
        dt.Columns.Add("Totale_Movimento")
        dt.Columns.Add("TotaleNetto_Doc")
        dt.Columns.Add("Bilancio")
        dt.Columns.Add("Data_Ora_Movimento")
        dt.Columns.Add("Id_Tipo_Pagamento")
        dt.Columns.Add("Id_Utente_Movimento")

        Dim rowCounter As Integer = 0
        Do Until fileContents.Length = 0
            Try
                rowCounter += 1
                Dim row As String = ""

                If InStr(fileContents, vbCrLf) Then
                    row = Strings.Left(fileContents, InStr(fileContents, vbCrLf) - 1)
                    fileContents = Strings.Right(fileContents, Len(fileContents) - InStr(fileContents, vbCrLf) - 1)
                Else
                    row = fileContents
                    fileContents = ""
                End If

                Dim dr As DataRow = dt.NewRow
                dr("Num_Doc") = CInt((Mid(row, 1, 10)))
                dr("Anno_Doc") = Mid(row, 11, 4)
                dr("Codice_Doc") = Trim(Mid(row, 15, 10))
                dr("Totale_Doc") = Mid(row, 25, 10)
                dr("Totale_Movimento") = Mid(row, 35, 10)
                dr("TotaleNetto_Doc") = Mid(row, 45, 10)
                dr("Bilancio") = Mid(row, 55, 10)
                dr("Data_Ora_Movimento") = Mid(row, 65, 8)
                dr("Id_Tipo_Pagamento") = Mid(row, 73, 10)
                dr("Id_Utente_Movimento") = Mid(row, 83, 10)
                dt.Rows.Add(dr)

            Catch ex As Exception
                MyOwnerForm.Cursor = Cursors.Default
                WaitingDialog.CloseDialog()
                Dim ErrorBody As String = "Errore nell'acquisizione del file prima nota" & rowCounter & vbCrLf &
                                ex.ToString
                MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                ErrorOccurred = True
                Exit Sub
            End Try
        Loop

        Dim dbContext As New HealthNET_DataEntities
        Dim cashFlowPaymentModeId_List As List(Of CashFlowPaymentMode) = dbContext.CashFlowPaymentModes.ToList


        Dim counter As Integer = 0
        WaitingDialog.ShowDialog("Importo dettagli prima nota ")
        For Each r In dt.Rows
            counter += 1

            Dim newCashFlow As CashFlow
            Dim invoiceNumber As Integer = CInt(r("Num_Doc"))
            Dim invoiceYear As Integer = CInt(r("Anno_Doc"))
            Dim invoiceDocCode As String = r("Codice_Doc")
            If invoiceDocCode.Trim = "M" Then
                invoiceDocCode = "O"
            End If

            Dim invoiceTotalAmount As Double = r("Totale_Doc")
            Dim cashFlowAmount As Double = r("Totale_Movimento")
            Dim invoiceNetAmount As Double = r("TotaleNetto_Doc")
            Dim Balance As Double = r("Bilancio")

            Dim cashFlowDateTime As DateTime = New Date(Mid(r("Data_Ora_Movimento"), 5, 4), Mid(r("Data_Ora_Movimento"), 3, 2), Mid(r("Data_Ora_Movimento"), 1, 2))
            Dim cashFlowPaymentModeId As String = r("Id_Tipo_Pagamento")
            If invoiceNumber = 1448 Then
                Dim popo = 2
            End If
            If cashFlowAmount <> 0 AndAlso (cashFlowPaymentModeId Is Nothing OrElse cashFlowPaymentModeId.Trim.Length = 0) Then
                WaitingDialog.CloseDialog()
                MyOwnerForm.Cursor = Cursors.Default
                Dim ErrorBody As String = "Errore nell'acquisizione del movimento di cassa per la fattura" & invoiceNumber & "/" & invoiceYear & "/" & invoiceDocCode & "/" & vbCrLf &
                                "Campo tipologia di pagamento non valorizzato"
                MessageBox.Show(ErrorBody, "Movimento non importato", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                ErrorOccurred = True
                Continue For
            End If


            Dim id_utente_movimento As String = "METH"

            If cashFlowPaymentModeId.Trim = "nn" Or cashFlowPaymentModeId.Trim = "" Or cashFlowAmount = 0 Then
                Continue For
            End If

            Dim relatedInvoice As Invoice = dbContext.Invoices.Where(Function(inv) inv.InvoiceNumber = invoiceNumber And
                                                                         inv.InvoiceYear = invoiceYear And
                                                                         inv.DocCode.Trim = invoiceDocCode.Trim).Single

            If relatedInvoice.IsCreditNote Then
                invoiceTotalAmount = -invoiceTotalAmount
            End If

            If relatedInvoice.Amount <> invoiceTotalAmount Then
                WaitingDialog.CloseDialog()
                MyOwnerForm.Cursor = Cursors.Default
                Dim ErrorBody As String = "Errore nell'acquisizione del movimento di cassa per la fattura" & invoiceNumber & "/" & invoiceYear & "/" & invoiceDocCode & "/" & vbCrLf &
                                "Totale documento incongruente"
                MessageBox.Show(ErrorBody, "Movimento non importato", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                ErrorOccurred = True
                Continue For
            End If


            If cashFlowPaymentModeId_List.Where(Function(pm) pm.CashFlowPaymentModeId = cashFlowPaymentModeId).Any = False Then
                WaitingDialog.CloseDialog()
                MyOwnerForm.Cursor = Cursors.Default
                Dim ErrorBody As String = "Errore nell'acquisizione del movimento di cassa per la fattura" & invoiceNumber & "/" & invoiceYear & "/" & invoiceDocCode & "/" & vbCrLf &
                                "Tipologia di pagamento non corretta: " & cashFlowPaymentModeId
                MessageBox.Show(ErrorBody, "Movimento non importato", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                ErrorOccurred = True
                Continue For
            End If

            Try
                Dim isUpdate As Boolean = False



                If relatedInvoice.CashFlows.Where(Function(cf) cf.CashFlowUser.Trim = "METH").Any Then
                    isUpdate = True
                    newCashFlow = relatedInvoice.CashFlows.Where(Function(cf) cf.CashFlowUser.Trim = "METH").FirstOrDefault
                Else
                    newCashFlow = New CashFlow
                End If

                newCashFlow.InvoiceNumber = invoiceNumber
                newCashFlow.InvoiceYear = invoiceYear
                newCashFlow.DocCode = invoiceDocCode
                newCashFlow.InvoiceTotal = invoiceTotalAmount

                If relatedInvoice.IsCreditNote Then
                    newCashFlow.InvoiceCashed = -cashFlowAmount
                Else
                    newCashFlow.InvoiceCashed = cashFlowAmount
                End If

                newCashFlow.Balance = Balance
                newCashFlow.DateTime = cashFlowDateTime
                newCashFlow.CashFlowPaymentModeId = cashFlowPaymentModeId
                newCashFlow.CashFlowUser = id_utente_movimento
                newCashFlow.PaidBy = "P"

                If isUpdate = False Then
                    dbContext.CashFlows.Add(newCashFlow)
                End If
            Catch ex As Exception
                WaitingDialog.CloseDialog()
                MyOwnerForm.Cursor = Cursors.Default
                Dim ErrorBody As String = "Errore nell'inserimento in db del movimento di cassa" & vbCrLf &
                                "per il documento" & invoiceNumber & "/" & invoiceYear & "/" & invoiceDocCode & vbCrLf &
                                ex.ToString
                MessageBox.Show(ErrorBody, "Movimento non importato", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                ErrorOccurred = True
                Exit Sub
            End Try
        Next
        WaitingDialog.CloseDialog()

        Try
            dbContext.SaveChanges()
        Catch ex As Exception
            WaitingDialog.CloseDialog()
            MyOwnerForm.Cursor = Cursors.Default
            Dim ErrorBody As String = "Errore nel salvataggio su db in Import Fatture emesse" & vbCrLf &
                ex.ToString
            MessageBox.Show(ErrorBody, "Importazione fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            GlobalFunctions.WriteImportFromHostLog(ErrorBody)
            ErrorOccurred = True
            Exit Sub
        End Try

    End Sub

    Private Sub ReverseTaxCode_Computing(TaxCode As String, ByRef Pat As Patient, patDbContext As HealthNET_DataEntities)
        Dim BirthDateString As String = Mid(TaxCode, 7, 5)
        Dim YearOffset1 As String = Strings.Left(Now.Year, 2)
        Dim YearOffset2 As String = Strings.Right(Now.Year, 2)

        Dim BirthYear As String = Mid(BirthDateString, 1, 2)

        If BirthYear < YearOffset2 Then
            BirthYear = YearOffset1 & BirthYear
        Else
            If BirthYear >= YearOffset2 Then BirthYear = YearOffset1 - 1 & BirthYear
        End If

        Dim BirthMonthLetter As String = Mid(BirthDateString, 3, 1)
        Dim BirthMonth As Integer

        Select Case BirthMonthLetter.ToUpper
            Case "A"
                BirthMonth = 1
            Case "B"
                BirthMonth = 2
            Case "C"
                BirthMonth = 3
            Case "D"
                BirthMonth = 4
            Case "E"
                BirthMonth = 5
            Case "H"
                BirthMonth = 6
            Case "L"
                BirthMonth = 7
            Case "M"
                BirthMonth = 8
            Case "P"
                BirthMonth = 9
            Case "R"
                BirthMonth = 10
            Case "S"
                BirthMonth = 11
            Case "T"
                BirthMonth = 12
        End Select

        Dim BirthDay As Integer = Mid(BirthDateString, 4, 2)
        If BirthDay >= 41 Then
            BirthDay -= 40
            Pat.SexId = 2
        Else
            Pat.SexId = 1
        End If


        Dim PatientBirthDate As New DateTime(BirthYear, BirthMonth, BirthDay)
        Pat.BirthDate = PatientBirthDate
        Dim intAge As Integer = ageComputing(PatientBirthDate)

        Dim PatientCityCode As String = Mid(TaxCode, 12, 4).ToUpper
        If PatientCityCode = "H5L1" Then
            PatientCityCode = "H501"
        End If
        Try

            Dim PatientCity As City = patDbContext.Cities.Where(Function(c) c.TaxCode = PatientCityCode).FirstOrDefault
            If PatientCity IsNot Nothing Then
                Pat.BirthCity = PatientCity.City1
                Pat.BirthCityCode = PatientCity.CodIstatCity
                Pat.CitizenshipCode = PatientCity.Citizenship.CountryCode
                Pat.Citizenship = PatientCity.Citizenship.CitizenshipDescription
                Pat.Area = PatientCity.Area
                Pat.AreaCode = PatientCity.AreaCode
            Else
                MyOwnerForm.Cursor = Cursors.Default
                WaitingDialog.CloseDialog()
                Dim ErrorBody As String = "Problema nella verifica del codice fiscale " & TaxCode & vbCrLf & vbCrLf &
                "Controllare la città di nascita con codice " & PatientCityCode & " desunta dal codice fiscale."
                ErrorBody &= vbCrLf & "Città di nascita non trovata"
                MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                GlobalFunctions.WriteImportFromHostLog(ErrorBody)
                Exit Sub
            End If
        Catch ex As Exception
            MyOwnerForm.Cursor = Cursors.Default
            WaitingDialog.CloseDialog()
            Dim ErrorBody As String = "Errore nella verifica del codice fiscale " & TaxCode & vbCrLf & vbCrLf &
                "Controllare la città di nascita" & vbCrLf & vbCrLf & ex.ToString
            ErrorBody &= vbCrLf & "Controllare la città di nascita"
            MessageBox.Show(ErrorBody, "Importazione Fallita", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            GlobalFunctions.WriteImportFromHostLog(ErrorBody)
            ErrorOccurred = True
            Exit Sub
        End Try

    End Sub

    Private Function ageComputing(BirthDate As DateTime) As Integer
        Dim intAge As Integer
        intAge = Today.Year - BirthDate.Year
        If (BirthDate > Today.AddYears(-intAge)) Then intAge -= 1
        Return intAge
    End Function

    Private Class RevenueCenters_Income
        Property RevenueCenterId As String
        Property NumberOfItems As Nullable(Of Integer)
        Property TotalIncome As Nullable(Of Double)
    End Class

End Module
