Attribute VB_Name = "Module2"
Sub CaigslistRenew()
Dim bot As New WebDriver
Dim keys As New SeleniumWrapper.keys
Dim values As String
Dim renewal As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
'Open Craigslist
bot.Start "chrome", "https://sandiego.craigslist.org/"
'Navigate To Login
bot.get "https://accounts.craigslist.org/login"
'Type Email
values = ActiveSheet.Range("D2").Value
bot.findElementByXPath("//*[@id=""inputEmailHandle""]").SendKeys (values)
'Type Password
values = ActiveSheet.Range("D4").Value
bot.findElementByXPath("//*[@id=""inputPassword""]").SendKeys (values)
'Press Enter To Login
bot.findElementByXPath("/html/body/section/section/div/div[1]/form/div[3]/button").Click
'Search for Renewals & Renew them if needed in first 4 posts
i = 1 'sets i = 1 becuase the craigslist links start at 1
j = 1
Do While i < 4 And j < 10
On Error GoTo errHandler
bot.findElementByXPath("//*[@id=""paginator""]/table/tbody/tr[" & j & "]/td[2]/div/form[4]/input[3]").Click
bot.findElementByXPath("//*[@id=""loginWidget""]/p[1]/strong/a").Click
i = i + 1
noError:
Loop
GoTo smoothExit
errHandler:
j = j + 1
If j = 10 Then
k = i - 1
i = 4
End If
Err.Clear
Resume noError
smoothExit:
bot.Close
bot.stop
MsgBox (k & " Ads were renewed")
End Sub
Public Sub findRenew()
i = 1 'sets i = 1 becuase the craigslist links start at 1
j = 1
Do While i < 4 Or j < 10
On Error GoTo errHandler
bot.findElementByXPath("//*[@id=""paginator""]/table/tbody/tr[1]/td[3]/a").Click
'bot.findElementByXPath("//*[@id=""paginator""]/table/tbody/tr[" & j & "]/td[2]/div/form[3]/input[3]").Click
'bot.findElementByXPath("//*[@id=""loginWidget""]/p[1]/strong/a").Click
i = i + 1
noError:
Loop
GoTo smoothExit
errHandler:

MessageBox.Show (Err.Number)

j = j + 1
If j = 10 Then
i = 4
End If
Err.Clear
Resume noError
smoothExit:
'i = 1
'Do While i < 5 Or j < 10
'bot.findElementByXPath("//*[@id=""paginator""]/table/tbody/tr[" & i & "]/td[2]/div/form[3]/input[3]").Click
'bot.findElementByXPath("//*[@id=""loginWidget""]/p[1]/strong/a").Click
'i = i + 1
'Loop
End Sub
Sub CraigslistRepost()
Dim bot As New WebDriver
Dim keys As New SeleniumWrapper.keys
Dim values As String
'Open Craigslist
bot.Start "chrome", "https://sandiego.craigslist.org/"
'Navigate To Login
bot.get "https://accounts.craigslist.org/login"
'Type Email
values = ActiveSheet.Range("D2").Value
bot.findElementByXPath("//*[@id=""inputEmailHandle""]").SendKeys (values)
'Type Password
values = ActiveSheet.Range("D4").Value
bot.findElementByXPath("//*[@id=""inputPassword""]").SendKeys (values)
'Press Enter To Login
bot.findElementByXPath("/html/body/section/section/div/div[1]/form/div[3]/button").Click
'Delete Most Recent Post
bot.findElementByXPath("//*[@id=""paginator""]/table/tbody/tr[1]/td[2]/div/form[2]/input[3]").Click
'Repost
bot.findElementByXPath("/html/body/article/section/div[1]/table/tbody/tr[2]/td[1]/div/form/input[2]").Click
'Continue
bot.findElementByXPath("//*[@id=""postingForm""]/div/button").Click
'Publish
bot.findElementByXPath("//*[@id=""publish_top""]/button").Click
'bot.Wait (10000)
'Continue
bot.findElementByXPath("//*[@id=""publish_top""]/button").Click
End Sub
