Attribute VB_Name = "Module1"
Sub Craigslist()
Dim bot As New WebDriver
Dim keys As New SeleniumWrapper.keys
Dim values As String
Dim ad As String
Dim location As String
ad = ActiveSheet.Range("A4")
Dim MSForms_DataObject As Object
Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
MSForms_DataObject.SetText ad
MSForms_DataObject.PutInClipboard
Set MSForms_DataObject = Nothing
'Open Craigslist
bot.Start "chrome", "https://sandiego.craigslist.org/"
'Navigate To Login
bot.get "https://sandiego.craigslist.org/"
'Login
bot.findElementByXPath("//*[@id=""postlks""]/li[2]/a").Click
'Type Email
values = ActiveSheet.Range("D2").Value
bot.findElementByXPath("//*[@id=""inputEmailHandle""]").SendKeys (values)
'Type Password
values = ActiveSheet.Range("D4").Value
bot.findElementByXPath("//*[@id=""inputPassword""]").SendKeys (values)
'Press Enter To Login
bot.findElementByXPath("/html/body/section/section/div/div[1]/form/div[3]/button").Click
'Navigate to Main Menu
bot.findElementByXPath("/html/body/article/header/a[1]").Click
'Post Add
bot.findElementByXPath("//*[@id=""post""]").Click
'Choose Housing Available
bot.findElementByXPath("/html/body/article/section/form/ul/li[4]/label/span[1]/input").Click
'Choose Apartments
bot.findElementByXPath("//*[@id=""picker""]/ul/li[2]/label/span[1]/input").Click
'Select Location
location = ActiveSheet.Range("C4").Value
bot.findElementByXPath(location).Click
'Choose Posting Title - Replace this code later with a form in excel to type in title before running
values = ActiveSheet.Range("A2").Value
bot.findElementByXPath("//*[@id=""PostingTitle""]").SendKeys (values)
'Choose Location
values = ActiveSheet.Range("B2").Value
bot.findElementByXPath("//*[@id=""GeographicArea""]").SendKeys (values)
'Enter Postal Code
values = ActiveSheet.Range("C2").Value
bot.findElementByXPath("//*[@id=""postal_code""]").SendKeys (values)
'Enter Ad Posting Body - Call Different Macro To open word, copy contents, Navigate back to chrome handle and call driver to use CTRL V
bot.findElementByXPath("//*[@id=""PostingBody""]").SendKeys keys.Control & "v"
'Enter Square Footage
bot.findElementByXPath("//*[@id=""Sqft""]").DoubleClick
values = ActiveSheet.Range("A6").Value
bot.findElementByXPath("//*[@id=""Sqft""]").SendKeys (values)
'Enter Rent Amount
values = ActiveSheet.Range("A8").Value
bot.findElementByXPath("//*[@id=""postingForm""]/div/div[4]/fieldset/label[2]/input").SendKeys (values)
'Select Bedrooms
values = ActiveSheet.Range("A10").Value
bot.findElementByXPath("//*[@id=""Bedrooms""]").SendKeys (values)
'Select Bathrooms
values = ActiveSheet.Range("A12").Value
bot.findElementByXPath("//*[@id=""bathrooms""]").SendKeys (values)
'Select Housing Type
values = ActiveSheet.Range("A14").Value
bot.findElementByXPath("//*[@id=""housing_type""]").SendKeys (values)
'Select Laundry
values = ActiveSheet.Range("A16").Value
bot.findElementByXPath("//*[@id=""laundry""]").SendKeys (values)
'Select Parking
values = ActiveSheet.Range("A18").Value
bot.findElementByXPath("//*[@id=""parking""]").SendKeys (values)
'Select Cats
bot.findElementByXPath("//*[@id=""pets_cat""]").Click
'Select Dogs
bot.findElementByXPath("//*[@id=""pets_dog""]").Click
'Select Non-Smoking
bot.findElementByXPath("//*[@id=""no_smoking""]").Click
'Select Phone
bot.findElementByXPath("//*[@id=""contact_phone_ok""]").Click
'Enter Phone Number
values = ActiveSheet.Range("A20").Value
bot.findElementByXPath("//*[@id=""contact_phone""]").SendKeys (values)
'Enter Street Address
values = ActiveSheet.Range("A22").Value
bot.findElementByXPath("//*[@id=""xstreet0""]").SendKeys (values)
'Enter City
values = ActiveSheet.Range("A24").Value
bot.findElementByXPath("//*[@id=""city""]").SendKeys (values)
'Enter State
values = ActiveSheet.Range("A26").Value
bot.findElementByXPath("//*[@id=""region""]").SendKeys (values)
'Continue
bot.findElementByXPath("//*[@id=""postingForm""]/div/button").Click
'Select Location Marker
'bot.findElementByXPath("//*[@id=""map""]/div[1]/div[4]/img").clickAndHold
'Continue
bot.findElementByXPath("//*[@id=""leafletForm""]/button[1]").Click
'Switch to Classic Photo Upload
bot.findElementByXPath("//*[@id=""classic""]").Click
'Open Images
values = ActiveSheet.Range("A28").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A29").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A30").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A31").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A32").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A33").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A34").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A35").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A36").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A37").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A38").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A39").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A40").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A41").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A42").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A43").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A44").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A45").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A46").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A47").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A48").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A49").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A50").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
bot.Wait (1000)
values = ActiveSheet.Range("A51").Value
bot.findElementByXPath("//*[@id=""uploader""]/form/input[3]").SendKeys (values)
'Complete Photo Upload
bot.findElementByXPath("/html/body/article/section/form/button").Click
End Sub
