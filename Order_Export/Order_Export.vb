Option Strict On

Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Net.Mail
Imports System.Xml

Module Order_Export
    Const CONNECTIONSTRING As String = "Database=magento;Data Source=localhost;" _
            & "User Id=root;Password=5BERA7279KL"
    Const KEYASCII As String = "1b3ccef66c6dad328afaea56b202b6d7"
    Dim successfulOrders As New ArrayList
    Dim mailTo, SMTPAddress As String

    Sub Main()
        readConfig()
        Dim fileTime As DateTime = DateTime.Now
        'Dim exportFileName As String = FILEPATH & "SERVU-" & fileTime.ToString("yyyyMMddHHmm") & ".txt"
        Dim exportFileName As String
        Dim orderConnection As New MySqlConnection(CONNECTIONSTRING)
        Dim billingAndShipping As New MySqlConnection(CONNECTIONSTRING)
        Dim shippingNumbers As New MySqlConnection(CONNECTIONSTRING)
        Dim currentOrder, previousOrder, parentProduct As Integer
        Dim exportOutput, couponValue, couponPrice As String
        Dim orderSQLStatement As String = ""
        Dim orderTypeSubject As String = ""
        Dim hasCoupon As Boolean = False

        For orderTypeIndex As Integer = 0 To 1
            If orderTypeIndex = 0 Then
                exportFileName = "SERVU-" & fileTime.ToString("yyyyMMddHHmm") & ".txt"
            ElseIf orderTypeIndex = 1 Then
                exportFileName = "SERVU-PAYPAL-" & fileTime.ToString("yyyyMMddHHmm") & ".txt"
            End If

            previousOrder = -1
            exportOutput = ""
            successfulOrders.Clear()

            Try
                orderConnection.Open()
                'processing
                'not paypal orders
                If orderTypeIndex = 0 Then
                    orderSQLStatement = "SELECT a.entity_id, a.coupon_code, a.base_discount_amount, " & _
                        "a.shipping_amount, a.customer_email, a.discount_description, a.customer_firstname,  a.customer_lastname, a.status, a.shipping_description, a.base_tax_amount, " & _
                        "b.cc_last4, b.cc_exp_month, b.cc_exp_year, b.cc_type, b.cc_number_enc, b.method, b.amount_paid, b.amount_ordered, b.additional_information, b.s4_uniqueid, c.sku, c.qty_ordered, " & _
                        "c.price, c.name, a.created_at, a.increment_id, a.customer_note, c.parent_item_id, d.exported FROM sales_flat_order AS a LEFT OUTER JOIN sales_flat_order_payment AS b on a.entity_id= b.parent_id " & _
                        "LEFT OUTER JOIN sales_flat_order_item AS c ON a.entity_id = c.order_id LEFT OUTER JOIN sales_flat_order_grid AS d on a.entity_id = d.entity_id " & _
                        "WHERE d.exported = 0 and b.method != 'paypal_standard' and a.status != 'canceled'" 'CHANGED: Used to read WHERE a.status='pending'
                    'paypal orders
                ElseIf orderTypeIndex = 1 Then
                    orderSQLStatement = "SELECT a.entity_id, a.coupon_code, a.base_discount_amount, " & _
                        "a.shipping_amount, a.customer_email, a.discount_description, a.customer_firstname,  a.customer_lastname, a.status, a.shipping_description, a.base_tax_amount, " & _
                        "b.last_trans_id, b.method, b.amount_paid, c.sku, c.qty_ordered, " & _
                        "c.price, c.name, a.created_at, a.increment_id, a.customer_note, c.parent_item_id, d.exported FROM sales_flat_order AS a LEFT OUTER JOIN sales_flat_order_payment AS b on a.entity_id= b.parent_id " & _
                        "LEFT OUTER JOIN sales_flat_order_item AS c ON a.entity_id = c.order_id LEFT OUTER JOIN sales_flat_order_grid AS d on a.entity_id = d.entity_id " & _
                        "WHERE d.exported = 0 and b.method = 'paypal_standard' and a.status = 'processing'" 'CHANGED: Used to read WHERE a.status='pending'
                End If

                Dim orderSQLCommand As MySqlCommand = New MySqlCommand(orderSQLStatement, orderConnection)
                Dim orderResultReader As MySqlDataReader = orderSQLCommand.ExecuteReader()

                While orderResultReader.Read()
                    'Create order header
                    currentOrder = orderResultReader.GetInt16(0)
                    parentProduct = -20

                    If currentOrder <> previousOrder Then
                        'this is a new order so we must create the order header
                        previousOrder = currentOrder
                        'remember to split out the address

                        Try
                            billingAndShipping.Open()
                            Dim billingAndShippingSQLStatment As String = "SELECT a.parent_id, a.fax, a.region, " & _
                                    "a.postcode, a.firstname, a.lastname, a.street, a.city, a.email, a.telephone, " & _
                                    "a.country_id, a.address_type, a.company, b.sourcecode FROM sales_flat_order_address AS a LEFT OUTER JOIN sales_flat_order as b ON a.parent_id = b.entity_id " & _
                                    "WHERE a.parent_id = " & currentOrder

                            Dim billingAndShippingSQLCommand As MySqlCommand = New MySqlCommand(billingAndShippingSQLStatment, billingAndShipping)
                            Dim billingAndShippingReader As MySqlDataReader = billingAndShippingSQLCommand.ExecuteReader()
                            Dim shippingInformation As New Dictionary(Of String, String)
                            Dim billingInformation As New Dictionary(Of String, String)

                            While billingAndShippingReader.Read()

                                If IfNull(billingAndShippingReader, "address_type", "") = "billing" Then
                                    billingInformation = CType(createCustomerInfo(billingAndShippingReader), Global.System.Collections.Generic.Dictionary(Of String, String))
                                Else
                                    shippingInformation = CType(createCustomerInfo(billingAndShippingReader), Global.System.Collections.Generic.Dictionary(Of String, String))
                                End If

                            End While

                            Dim orderDateString As String = IfNull(orderResultReader, "created_at", "01/01/2012")
                            Dim orderDate As DateTime = Convert.ToDateTime(orderDateString)
                            Dim formatedDate As String = orderDate.ToString("MM/dd/yy")
                            Dim crypt As New Chilkat.Crypt2()
                            Dim expMonth, expYear, exp, couponCode, creditCardField, shippingDescription As String

                            exp = ""
                            couponValue = ""
                            couponCode = ""
                            creditCardField = ""
                            crypt = CType(decryptCC(), Chilkat.Crypt2)

                            exportOutput &= "A" & CType(createField(billingInformation("telephone"), 28), String)
                            exportOutput &= CType(createField(billingInformation("company"), 40), String)
                            exportOutput &= CType(createField(billingInformation("firstName"), 12), String)
                            exportOutput &= CType(createField(billingInformation("lastName"), 17), String)
                            exportOutput &= CType(createField(billingInformation("addressOne"), 35), String)
                            exportOutput &= CType(createField(billingInformation("addressTwo"), 35), String)
                            'City field has 3 other filler fields with it, see solutions documentation
                            exportOutput &= CType(createField(billingInformation("city"), 46), String)
                            exportOutput &= CType(createField(billingInformation("state"), 25), String)
                            exportOutput &= CType(createField(billingInformation("country"), 25), String)
                            exportOutput &= CType(createField(billingInformation("zip"), 10), String)
                            exportOutput &= CType(createField(billingInformation("sourcecode"), 8), String)
                            '#TODO: Orderdate also uses ck bank name(40), Bank City(30), Ck Num(6), ck Account Num(53)
                            exportOutput &= CType(createField(formatedDate, 137), String)
                            'Fax also uses fax_additional(10), srvc bur commnt(65),  dflt_csr_id(4), CUST_TYPE(8)
                            'batch(4), ctsy_title(10)
                            exportOutput &= CType(createField(billingInformation("fax"), 111), String)
                            '#TODO: Commercial field to be implemented later Values are "C" and "R" for commercial
                            'and residential. Length is 1 the actual values are 1 for commercial and 2 for residential
                            exportOutput &= "1" & vbCrLf
                            '00.00 is the Discount_perc field with a length of 4
                            exportOutput &= "B" & "00.00"
                            'VI, MC, DIS, AX, CHK, COD, PPL
                            exportOutput &= CType(createField(CStr(determinePayment(orderResultReader)), 4), String)

                            'decrypt the credit card number field if there is a credit card #, then do a string replace for Solution
                            If orderTypeIndex = 0 Then
                                If IfNull(orderResultReader, "cc_number_enc", "") <> "" Then
                                    creditCardField = crypt.DecryptStringENC(orderResultReader.GetString("cc_number_enc"))

                                    Dim ccLength As Integer = Len(Trim(creditCardField))
                                    Dim leftPortion As String = Microsoft.VisualBasic.Left(creditCardField, 6)
                                    Dim rightPortion As String = Microsoft.VisualBasic.Right(creditCardField, 4)
                                    Dim middlePortion As String = ""
                                    ccLength = ccLength - 10

                                    For index As Integer = 1 To ccLength
                                        middlePortion &= "X"
                                    Next

                                    creditCardField = leftPortion & middlePortion & rightPortion

                                End If
                            End If

                            exportOutput &= CType(createField(creditCardField, 17), String)

                            If orderTypeIndex = 0 Then
                                expMonth = IfNull(orderResultReader, "cc_exp_month", "")
                            Else
                                expMonth = ""
                            End If

                            'create the expiration date for credit cards
                            If expMonth <> "" And expMonth.Length < 3 And creditCardField <> "" And orderTypeIndex = 0 Then
                                expMonth = expMonth.PadLeft(2, CChar("0"))

                                expYear = IfNull(orderResultReader, "cc_exp_year", "")
                                'If expYear <> "" Then
                                'expYear = expYear.Substring(2, 2)
                                'End If

                                exp = expMonth & expYear

                            End If

                            'Expiration also contains approval code(8), customer folloup(1) Y=Customer service followup
                            'or catalog request No record type = C or D, and installment flag(1)Y or N
                            exportOutput &= CType(createField(exp, 13), String)

                            'Contains Defered Billing(1) and  AUTH_DATE(8)
                            exportOutput &= CType(createField(orderResultReader, "customer_email", "", 35), String)
                            'AUTH_DATE
                            If CStr(determinePayment(orderResultReader)) <> "CHK" Then
                                exportOutput &= CType(createField(formatedDate, 8), String)
                            Else
                                exportOutput &= CType(createField("", 8), String)
                            End If

                            'Create field for paypal and shift 4 amounts here, AUTH_AMOUNT(11) field
                            If CStr(determinePayment(orderResultReader)) <> "CHK" And orderTypeIndex = 0 Then
                                Dim shift4price = CType(createField(orderResultReader, "amount_ordered", "00000000.00", 11), String)
                                shift4price = Left(shift4price, shift4price.IndexOf(".") + 3)
                                shift4price = shift4price.PadLeft(11, CChar("0"))
                                exportOutput &= shift4price
                            ElseIf orderTypeIndex = 1 Then
                                Dim paypalprice = CType(createField(orderResultReader, "amount_paid", "00000000.00", 11), String)
                                paypalprice = Left(paypalprice, paypalprice.IndexOf(".") + 3)
                                paypalprice = paypalprice.PadLeft(11, CChar("0"))
                                exportOutput &= paypalprice
                            Else
                                exportOutput &= CType(createField("00000000.00", 11), String)
                            End If

                            

                            'exportOutput &= CType(createField(billingInformation("firstName") & " " & billingInformation("lastName"), 35), String)
                            'exportOutput &= CType(createField(billingInformation("addressOne"), 35), String)
                            'exportOutput &= CType(createField(billingInformation("zip"), 10), String)
                            'This next field encompasses Invoice Number(10)'
                            exportOutput &= CType(createField(" ", 10), String)
                            exportOutput &= CType(createField(orderResultReader, "increment_id", "NOORDER#", 19, "right"), String)

                            'PAYPAL ID
                            If orderTypeIndex = 1 Then
                                exportOutput &= CType(createField(orderResultReader.GetString("last_trans_id"), 35), String)
                            Else
                                exportOutput &= CType(createField("", 35), String)
                            End If

                            'Get the credit card token for shift 4 and create the token expiration date
                            If orderTypeIndex = 0 And CStr(determinePayment(orderResultReader)) <> "CHK" Then
                                exportOutput &= CType(createField(orderResultReader, "s4_uniqueid", "", 20), String)

                                Dim s4ExpDate As DateTime = Convert.ToDateTime(orderDateString)
                                s4ExpDate = s4ExpDate.AddYears(2)
                                exportOutput &= CType(createField(s4ExpDate.ToString("MM/dd/yy"), 10), String)
                            Else
                                exportOutput &= CType(createField("", 30), String)
                            End If

                            shippingNumbers.Open()
                            Dim shippingNumbersSQLStatment As String = "SELECT a.value FROM servu_shipping_misc_information_order AS a " & _
                                    "WHERE a.order_id = " & currentOrder
                            Dim shippingNumbersSQLCommand As MySqlCommand = New MySqlCommand(shippingNumbersSQLStatment, shippingNumbers)
                            Dim shippingNumbersReader As MySqlDataReader = shippingNumbersSQLCommand.ExecuteReader()
                            Dim confirmationNumbers As String = ""
                            Dim shippingOptions As String = ""
                            Dim shippingMiscArray As New ArrayList

                            While shippingNumbersReader.Read()
                                shippingMiscArray.Add(IfNull(shippingNumbersReader, "value", ""))
                            End While

                            For Each shippingParts As String In shippingMiscArray
                                If shippingParts = "s_extra_lfg" Then
                                    shippingOptions &= " Lift Gate Requested "
                                ElseIf shippingParts = "s_extra_notify" Then
                                    shippingOptions &= " Pre-Notify Requested "
                                ElseIf shippingParts = "s_extra_residence" Then
                                    shippingOptions &= " Residence Delivery "
                                Else
                                    confirmationNumbers &= " " & shippingParts
                                End If
                            Next

                            shippingNumbers.Close()

                            If (confirmationNumbers <> "") Then
                                Dim confirmationString = "Conway Confirmation Numbers " & confirmationNumbers
                                confirmationNumbers = confirmationString
                            End If

                            If IfNull(orderResultReader, "coupon_code", "") <> "" Then
                                couponCode = "COUPON CODE: " & IfNull(orderResultReader, "coupon_code", "")
                            End If

                            If IfNull(orderResultReader, "base_discount_amount", "0.0000") <> "0.0000" Then
                                couponPrice = IfNull(orderResultReader, "base_discount_amount", "")
                                couponValue = "COUPON value: " & IfNull(orderResultReader, "base_discount_amount", "")
                                hasCoupon = True
                            End If

                            shippingDescription = IfNull(orderResultReader, "shipping_description", "")

                            exportOutput &= CType(createField("TAX Amount: " & IfNull(orderResultReader, "base_tax_amount", "") & " " & IfNull(orderResultReader, "customer_note", "") & " " & couponCode & " " & couponValue & " " & shippingDescription & " " & confirmationNumbers & " " & shippingOptions, 750), String) & vbCrLf
                            exportOutput &= "C" & CType(createField(shippingInformation("telephone"), 10), String)
                            exportOutput &= CType(createField(shippingInformation("company"), 40), String)
                            exportOutput &= CType(createField(shippingInformation("firstName"), 12), String)
                            exportOutput &= CType(createField(shippingInformation("lastName"), 17), String)
                            exportOutput &= CType(createField(shippingInformation("addressOne"), 35), String)
                            exportOutput &= CType(createField(shippingInformation("addressTwo"), 35), String)
                            exportOutput &= CType(createField(shippingInformation("city"), 46), String)
                            exportOutput &= CType(createField(shippingInformation("state"), 25), String)
                            exportOutput &= CType(createField(shippingInformation("country"), 25), String)
                            exportOutput &= CType(createField(shippingInformation("zip"), 10), String)
                            'This is tax rate
                            exportOutput &= CType(createField("00.00", 5), String)
                            exportOutput &= CType(createField(orderResultReader, "shipping_amount", "000.00", 6), String)
                            'This is misc amt
                            exportOutput &= CType(createField("00000.00", 8), String)
                            'This is handling amt
                            exportOutput &= CType(createField("000.00", 6), String)
                            'Gift Message
                            exportOutput &= CType(createField("", 179), String)
                            'PO Number from Web Number
                            exportOutput &= CType(createField(orderResultReader, "increment_id", "", 15, "right"), String)
                            'Ctsy Title
                            exportOutput &= CType(createField("", 10), String)
                            exportOutput &= CType(createField(orderResultReader, "customer_email", "", 35), String)
                            'Print hard copy
                            exportOutput &= CType(createField("Y", 1), String) & vbCrLf

                            If IfNull(orderResultReader, "parent_item_id", "NULL") = "NULL" Then
                                exportOutput &= CType(buildLineItem(orderResultReader), String)
                            End If

                            billingAndShippingReader.Close()
                            successfulOrders.Add(orderResultReader.GetInt32("entity_id"))
                        Catch e As MySqlException
                            LogError(DateTime.Now & "Error in billing connection/query: " & e.ToString())
                        Finally
                            billingAndShipping.Close()
                        End Try

                    Else
                        'this is still the same order so we simply need to append a new line item
                        If IfNull(orderResultReader, "parent_item_id", "NULL") = "NULL" Then
                            exportOutput &= CType(buildLineItem(orderResultReader), String)
                        End If

                    End If

                End While

                If hasCoupon Then
                    exportOutput &= CType(buildFakeLineItem(couponPrice), String) & vbCrLf
                End If

                orderResultReader.Close()

            Catch e As MySqlException
                LogError(DateTime.Now & "Error in order connection/query: " & e.ToString())
            Finally
                orderConnection.Close()
            End Try

            If orderTypeIndex = 0 Then
                orderTypeSubject = "Normal Order Export: "
            ElseIf orderTypeIndex = 1 Then
                orderTypeSubject = "Paypal Order Export: "
            End If

            If successfulOrders.Count <> 0 Then
                Console.Write(exportOutput)
                sendExportResults(orderTypeSubject & exportFileName, exportFileName & " exported successfully.", exportFileName, exportOutput)
                'The following lines have been commented out for testing
                'updateOrders()
            Else
                sendExportResults(orderTypeSubject & "No orders", exportFileName & " There were no orders to export.", exportFileName, "")
            End If
        Next
        Threading.Thread.Sleep(100000)

    End Sub
    Sub readConfig()
        SMTPAddress = "192.168.1.50"
        mailTo = "travishill@servu-online.com"
        Try
            Dim doc As New System.Xml.XmlDocument
            doc.Load(My.Application.Info.DirectoryPath & "\config.xml")
            Dim list = doc.GetElementsByTagName("name")

            If (doc.GetElementsByTagName("SMTP")(0).InnerText <> "") Then
                SMTPAddress = doc.GetElementsByTagName("SMTP")(0).InnerText
            End If


            If (doc.GetElementsByTagName("MailTo")(0).InnerText <> "") Then
                mailTo = doc.GetElementsByTagName("MailTo")(0).InnerText
            End If
        Catch e As Exception
            LogError("Error in xml: " & e.ToString())
        End Try

    End Sub
    Sub sendExportResults(ByVal subject As String, ByVal body As String, ByVal exportFileName As String, ByVal exportFile As String)
        Dim mail As New MailMessage()
        Dim smtp As New SmtpClient(SMTPAddress)

        mail.From = New MailAddress(mailTo)
        mail.To.Add(mailTo)

        mail.Subject = subject
        mail.Body = body   

        If exportFile <> "" Then
            Dim ms As MemoryStream = New MemoryStream()
            Dim sw As StreamWriter = New StreamWriter(ms)

            'sw = My.Computer.FileSystem.OpenTextFileWriter("c:\test.txt", True)
            'sw.WriteLine(exportFile)
            'sw.Close()

            sw.WriteLine(exportFile)
            sw.Flush()
            ms.Position = 0
            mail.Attachments.Add(New Attachment(ms, exportFileName, "text/plain"))
        End If

        smtp.Send(mail)
    End Sub

    Private Sub updateOrders()
        Dim conn As New MySqlConnection(CONNECTIONSTRING)
        Dim cmd As New MySqlCommand()
        Dim transaction As MySqlTransaction
        Dim whereGridClause, sqlStatement As String

        'whereClause = String.Join(" OR sales_flat_order.entity_id = ", successfulOrders.ToArray())
        whereGridClause = String.Join(" OR sales_flat_order_grid.entity_id = ", successfulOrders.ToArray())

        sqlStatement = "UPDATE sales_flat_order JOIN sales_flat_order_grid ON " & _
        "sales_flat_order_grid.entity_id = sales_flat_order.entity_id " & _
        "SET sales_flat_order.exported = 1, sales_flat_order_grid.exported = 1 " & _
        "WHERE sales_flat_order.entity_id = " & whereGridClause
        'processing, fulfilling
        'sqlStatement = "UPDATE sales_flat_order_grid SET sales_flat_order_grid.exported = 1 WHERE sales_flat_order_grid.entity_id = " & whereGridClause & ";"
        'fulfilling
        'sqlStatement &= "UPDATE sales_flat_order_grid SET sales_flat_order_grid.status = 'processing' WHERE sales_flat_order_grid.entity_id = " & whereGridClause & ";"

        Try
            conn.Open()
            transaction = conn.BeginTransaction()

            cmd.Connection = conn
            cmd.Transaction = transaction

            cmd.CommandText = sqlStatement
            cmd.ExecuteNonQuery()

            transaction.Commit()
            conn.Close()

        Catch e As MySqlException
            transaction.Rollback()
            LogError(DateTime.Now & "Error in commiting update: " & e.ToString())
        End Try

    End Sub

    Private Function determinePayment(ByVal dr As MySqlDataReader) As String

        Dim payment, returnValue As String

        returnValue = IfNull(dr, "method", "")

        If returnValue = "ccsave" Or returnValue = "shift4payment" Then
            payment = IfNull(dr, "cc_type", "")
            If payment = "AE" Then
                payment = "AX"
            ElseIf payment = "DI" Then
                payment = "DIS"
            ElseIf payment = "VS" Then
                payment = "VI"
            End If
        ElseIf returnValue = "checkmo" Then
            payment = "CHK"
        ElseIf returnValue = "paypal_standard" Then
            payment = "PPL"
        Else
            payment = "NOPA"
        End If

        createField(returnValue, 4)

        Return payment

    End Function

    Private Function buildLineItem(ByVal dr As MySqlDataReader) As String
        'This does not properly split out the sku numbers for attributes, if the sku has just an
        '^ then it is fine, but anything with an = or a mixture of = and ^ will result in the 
        'entire sku being placed into the field for example WAYU-1275=Wood=Antiq  will show, but
        'the comments area will also have the proper splits but with the sku number.

        Dim lineItem, qtyOrdered, price, sku, comments, previousPart As String
        Dim importableAttributes As New ArrayList
        comments = ""
        lineItem = ""
        previousPart = "#FIRSTPART#"

        sku = IfNull(dr, "sku", "BADNUMBER")

        Dim splitSku() As String = Regex.Split(sku, "([=^#*])")

        For Each skuPart As String In splitSku
            If previousPart = "#FIRSTPART#" Then
                lineItem = "D" & CType(createField(skuPart, 20), String)
            ElseIf previousPart = Chr(61) Then
                comments &= skuPart & " "
            ElseIf previousPart = Chr(94) Then
                importableAttributes.Add(skuPart)
            End If
            previousPart = skuPart
        Next


        '#TODO: There are two attribute fields that need to be ripped from the current item
        'The two lines of code that follow are simply place holders until this is established in Magento
        If importableAttributes.Count >= 1 Then
            lineItem &= CType(createField(CStr(importableAttributes(0)), 10), String)

            If importableAttributes.Count >= 2 Then
                lineItem &= CType(createField(CStr(importableAttributes(1)), 10), String)
            Else
                lineItem &= CType(createField("", 10), String)
            End If
        Else
            lineItem &= CType(createField("", 20), String)
        End If

        qtyOrdered = CType(createField(dr, "qty_ordered", "0.000", 5), String)
        qtyOrdered = Left(qtyOrdered, qtyOrdered.IndexOf("."))
        lineItem &= CType(createField(qtyOrdered, 5), String)
        price = CType(createField(dr, "price", "0000.0000", 8), String)
        price = Left(price, price.IndexOf(".") + 3)
        lineItem &= price.PadLeft(7, CChar("0"))
        'Gift Wrap
        lineItem &= CType(createField("N", 1), String)
        'Gift wrap fee
        lineItem &= CType(createField("00.00", 5), String)
        'Giftwrap message, contains variable_kit(1), attribute 3(10), attribute 4(10), attribute 5(10)
        lineItem &= CType(createField(comments, 320), String)
        lineItem &= CType(createField(" ", 1), String)
        'This sucks, you know it, I know it.
        If importableAttributes.Count >= 3 Then
            lineItem &= CType(createField(CStr(importableAttributes(2)), 10), String)

            If importableAttributes.Count >= 4 Then
                lineItem &= CType(createField(CStr(importableAttributes(3)), 10), String)

                If importableAttributes.Count = 5 Then
                    lineItem &= CType(createField(CStr(importableAttributes(4)), 10), String)
                Else
                    lineItem &= CType(createField("", 10), String)
                End If
            Else
                lineItem &= CType(createField("", 20), String)
            End If
        Else
            lineItem &= CType(createField("", 30), String)
        End If
        lineItem &= CType(createField(dr, "name", "NODESCRIPTION", 74), String) & vbCrLf

        Return lineItem

    End Function

    Private Function buildFakeLineItem(ByVal couponValue As String) As String
        Dim lineItem, price, sku, comments As String
        comments = ""
        lineItem = ""

        sku = "COUPON5"

        lineItem = "D" & CType(createField(sku, 20), String)

        lineItem &= CType(createField("", 20), String)
        lineItem &= CType(createField(CStr(1), 5), String)
        price = couponValue
        price = Left(price, price.IndexOf(".") + 3)
        lineItem &= price.PadLeft(7, CChar("0"))
        'Gift Wrap
        lineItem &= CType(createField("N", 1), String)
        'Gift wrap fee
        lineItem &= CType(createField("00.00", 5), String)
        'Giftwrap message, contains variable_kit(1), attribute 3(10), attribute 4(10), attribute 5(10)
        lineItem &= CType(createField(comments, 320), String)
        lineItem &= CType(createField(" ", 31), String)
        lineItem &= CType(createField("REVIEW DISCOUNT COUPON", 74), String) & vbCrLf

        Return lineItem
    End Function

    Private Function createCustomerInfo(ByVal dr As MySqlDataReader) As System.Collections.Generic.Dictionary(Of String, String)
        'Create array for shipping and billing information
        Dim returnValues As New Dictionary(Of String, String)
        Dim address() As String
        Dim state As String

        state = IfNull(dr, "region", "")

        If Len(state) > 2 Then
            Select Case state
                Case "Alabama"
                    state = "AL"
                Case "Alaska"
                    state = "AK"
                Case "Arizona"
                    state = "AZ"
                Case "Arkansas"
                    state = "AR"
                Case "California"
                    state = "CA"
                Case "Colorado"
                    state = "CO"
                Case "Connecticut"
                    state = "CT"
                Case "Delaware"
                    state = "DE"
                Case "Florida"
                    state = "FL"
                Case "Georgia"
                    state = "GA"
                Case "Hawaii"
                    state = "HI"
                Case "Idaho"
                    state = "ID"
                Case "Illinois"
                    state = "IL"
                Case "Indiana"
                    state = "IN"
                Case "Iowa"
                    state = "IA"
                Case "Kansas"
                    state = "KS"
                Case "Kentucky"
                    state = "KY"
                Case "Louisiana"
                    state = "LA"
                Case "Maine"
                    state = "ME"
                Case "Maryland"
                    state = "MD"
                Case "Massachusetts"
                    state = "MA"
                Case "Michigan"
                    state = "MI"
                Case "Minnesota"
                    state = "MN"
                Case "Mississippi"
                    state = "MS"
                Case "Missouri"
                    state = "MO"
                Case "Montana"
                    state = "MT"
                Case "Nebraska"
                    state = "NE"
                Case "Nevada"
                    state = "NV"
                Case "New Hampshire"
                    state = "NH"
                Case "New Jersey"
                    state = "NJ"
                Case "New Mexico"
                    state = "NM"
                Case "New York"
                    state = "NY"
                Case "North Carolina"
                    state = "NC"
                Case "North Dakota"
                    state = "ND"
                Case "Ohio"
                    state = "OH"
                Case "Oklahoma"
                    state = "OK"
                Case "Oregon"
                    state = "OR"
                Case "Pennsylvania"
                    state = "PA"
                Case "Rhode Island"
                    state = "RI"
                Case "South Carolina"
                    state = "SC"
                Case "South Dakota"
                    state = "SD"
                Case "Tennessee"
                    state = "TN"
                Case "Texas"
                    state = "TX"
                Case "Utah"
                    state = "UT"
                Case "Vermont"
                    state = "VT"
                Case "Virginia"
                    state = "VA"
                Case "Washington"
                    state = "WA"
                Case "West Virginia"
                    state = "WV"
                Case "Wisconsin"
                    state = "WI"
                Case "Wyoming"
                    state = "WY"
                Case "Armed Forces Africa"
                Case "Armed Forces Canada"
                Case "Armed Forces Europe"
                Case "Armed Forces Middle East"
                    state = "AE"
                Case "Armed Forces Americas"
                    state = "AA"
                Case "Armed Forces Canada"
                    state = "AE"
            End Select
        End If

        returnValues("firstName") = IfNull(dr, "firstname", "")
        returnValues("lastName") = IfNull(dr, "lastname", "")
        returnValues("city") = IfNull(dr, "city", "")
        returnValues("state") = state
        returnValues("zip") = IfNull(dr, "postcode", "")
        If IfNull(dr, "address_type", "") = "billing" Then
            returnValues("sourcecode") = IfNull(dr, "sourcecode", "")
        End If
        returnValues("country") = IfNull(dr, "country_id", "")
        If (returnValues("country")) = "US" Then
            returnValues("country") = "USA"
        End If
        returnValues("email") = IfNull(dr, "email", "")
        returnValues("telephone") = Regex.Replace(IfNull(dr, "telephone", ""), "[^0-9]", "")
        returnValues("fax") = IfNull(dr, "fax", "")
        returnValues("company") = IfNull(dr, "company", "")

        address = CType(separateAddress(IfNull(dr, "street", "")), String())

        returnValues("addressOne") = address(0)
        returnValues("addressTwo") = address(1)

        Return returnValues

    End Function

    Private Function separateAddress(ByVal address As String) As String()
        ' Split string based on new line
        Dim parts() As String = address.Split(Chr(10), Chr(13))

        If parts.Length = 1 Then
            Dim preserved = parts(0)
            ReDim parts(1)
            parts(0) = preserved
            parts(1) = ""
        End If

        Return parts

    End Function

    Private Function decryptCC() As Chilkat.Crypt2
        Dim crypt As New Chilkat.Crypt2()
        Dim success As Boolean
        Dim ivAscii As String = "abcdefgh"

        success = crypt.UnlockComponent("SERVCOCrypt_BE4z29b0WVIh")

        crypt.CryptAlgorithm = "blowfish2"
        crypt.CipherMode = "ecb"
        crypt.KeyLength = 256
        crypt.PaddingScheme = 3
        crypt.EncodingMode = "base64"
        crypt.Charset = "utf-8"

        crypt.SetEncodedIV(ivAscii, "ascii")
        'key ascii is the magento key
        crypt.SetEncodedKey(KEYASCII, "ascii")

        Return crypt
    End Function

    Private Sub LogError(ByVal e As String)
        'Utility function to log errors
        Dim fileName As String = My.Application.Info.DirectoryPath & "\logs\log.txt"

        If File.Exists(fileName) Then
            Using fileWriter As StreamWriter = New StreamWriter(fileName, True)
                fileWriter.Write(e)
            End Using
        Else
            Using fileWriter As StreamWriter = New StreamWriter(fileName)
                fileWriter.Write(e)
            End Using
        End If

    End Sub

    Public Function IfNull(Of T)(ByVal dr As MySqlDataReader, ByVal fieldName As String, ByVal _default As T) As T

        If IsDBNull(dr(fieldName)) Then
            Return _default
        Else
            Return CType(dr(fieldName), T)
        End If

    End Function

    Private Function createField(ByVal fieldValue As String, ByVal totalSpaces As Integer) As String

        If fieldValue.Length > totalSpaces Then
            Return fieldValue.Substring(0, totalSpaces)
        Else
            Return fieldValue.PadRight(totalSpaces, CChar(" "))
        End If

    End Function

    Private Function createField(Of T)(ByVal dr As MySqlDataReader, ByVal fieldName As String, ByVal _default As T, _
                                    ByVal totalSpaces As Integer, Optional ByVal padDirection As String = "") As String
        Dim fieldValue As String

        fieldValue = IfNull(dr, fieldName, _default).ToString

        If fieldValue.Length > totalSpaces Then
            Return fieldValue.Substring(0, totalSpaces)
        Else
            If padDirection = "left" Then
                Return fieldValue.PadLeft(totalSpaces, CChar(" "))
            Else
                Return fieldValue.PadRight(totalSpaces, CChar(" "))
            End If
        End If

    End Function

End Module
