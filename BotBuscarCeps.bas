Attribute VB_Name = "Módulo1"
Sub BuscaCep()

On Error Resume Next 'Se houver algum erro o código continuará seu fluxo normal

Dim bot As New WebDriver
Dim url As String
Dim element As WebElement
Dim optionElement As Object
Dim htmlAnswer As String
Dim CEPs() As String

url = "https://www2.correios.com.br/sistemas/buscacep/buscaFaixaCep.cfm"
bot.Start "edge", "C:\Users\SEU-USUÁRIO-AQUI\AppData\Local\SeleniumBasic\edgedriver.exe"


For i = 2 To 5571
    bot.Get url
    Set element = bot.FindElementByName("UF")
    Set optionElement = element.FindElementsByTag("option")
    For j = 1 To optionElement.Count
        If optionElement.Item(j).Attribute("value") = Cells(i, 1) Then
            optionElement.Item(j).Click
            Exit For
        End If
    Next j
    Set element = bot.FindElementByName("Localidade")
    element.SendKeys Cells(i, 5)
    Set element = bot.FindElementByClass("btn2")
    element.Click
    Set element = bot.FindElementByXPath("/html/body/div[1]/div[3]/div[2]/div/div/div[2]/div[2]/div[2]/table[2]/tbody/tr[3]/td[2]")
    htmlAnswer = element.Text
    CEPs = Split(htmlAnswer, " a ")
    If Cells(i - 1, 2).Value = CEPs(0) And Cells(i - 1, 3).Value = CEPs(1) Then
        Cells(i, 2).Value = "Erro" 'Controle para prencher com "Erro" as cidades que não foi possível buscar o CEP
        Cells(i, 3).Value = "Erro"
    Else
        Cells(i, 2).Value = CEPs(0)
        Cells(i, 3).Value = CEPs(1)
    End If
Next i
bot.Quit

End Sub

