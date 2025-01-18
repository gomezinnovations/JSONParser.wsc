# JSON Parser Windows Script Component (WSC)

A **lightweight, high-performance** JSON parser for **Classic ASP and Windows Script Host (WSH)**, enabling **JSON parsing (`ParseJson`)** and **serialization (`Stringify`)** without third-party dependencies.

## 🚀 Features
✅ **Convert JSON to Objects** (`ParseJson`)  
✅ **Convert Objects to JSON** (`Stringify`)  
✅ **Supports Nested Objects & Arrays**  
✅ **Handles Numbers, Booleans, and Nulls Correctly**  
✅ **Graceful Error Handling with JSON-Formatted Errors**  
✅ **Works in Classic ASP, VBScript, and Any COM-Compatible System**  

---

## 📂 Installation
### Step 1: Copy the WSC File
Save `JSONParser.wsc` to a **trusted location** on your server (e.g., `C:\Windows\System32\` or your project directory).

### Step 2: Register the Component
Open **Command Prompt as Administrator** and run:
```cmd
regsvr32 "C:\path\to\JSONParser.wsc"
```
📌 Replace `"C:\path\to\JSONParser.wsc"` with the actual path.

### Step 3: Verify Registration
Run the following command:
```cmd
cscript //nologo
```
Then, type:
```vbscript
Set jsonParser = CreateObject("JSON.Parser")
WScript.Echo TypeName(jsonParser)
```
If successful, it should return `"Object"`.

---

## 📝 Usage
### 🔹 Classic ASP Example
```asp
<%
Dim jsonParser, parsedJson, jsonData

' Create an instance of the JSON parser
Set jsonParser = Server.CreateObject("JSON.Parser")

' Sample JSON string
jsonData = "{""name"":""John Doe"",""age"":30,""skills"":[""ASP"",""VBScript"",""SQL""],""address"":{""city"":""NYC""}}"

' Parse JSON
Set parsedJson = jsonParser.ParseJson(jsonData)

' Check for errors
If parsedJson.Exists("error") Then
    Response.Write "<b>Error:</b> " & parsedJson("error") & "<br>"
Else
    Response.Write "<b>Name:</b> " & parsedJson("name") & "<br>"
    Response.Write "<b>City:</b> " & parsedJson("address")("city") & "<br>"
End If

' Cleanup
Set jsonParser = Nothing
Set parsedJson = Nothing
%>
```

### 🔹 VBScript Example
```vbscript
Dim jsonParser, jsonData, parsedJson

' Create JSON parser instance
Set jsonParser = CreateObject("JSON.Parser")

' JSON String
jsonData = "{""message"":""Hello, World!"",""success"":true}"

' Parse JSON
Set parsedJson = jsonParser.ParseJson(jsonData)

' Output
If parsedJson.Exists("error") Then
    WScript.Echo "Error: " & parsedJson("error")
Else
    WScript.Echo "Message: " & parsedJson("message")
End If

' Cleanup
Set jsonParser = Nothing
Set parsedJson = Nothing
```

---

## 🔄 Stringify JSON Example
Convert a **Dictionary object** back to JSON.

### 🔹 Classic ASP / VBScript
```vbscript
Dim jsonParser, objDict, jsonString

' Create instance
Set jsonParser = CreateObject("JSON.Parser")

' Create a JSON object (Dictionary)
Set objDict = CreateObject("Scripting.Dictionary")
objDict.Add "name", "Jane Doe"
objDict.Add "age", 25
objDict.Add "city", "Los Angeles"

' Convert to JSON string
jsonString = jsonParser.Stringify(objDict)

' Output JSON
WScript.Echo jsonString

' Cleanup
Set jsonParser = Nothing
Set objDict = Nothing
```
**Expected Output:**
```json
{"name":"Jane Doe","age":25,"city":"Los Angeles"}
```

---

## ❗ Error Handling
The parser **never crashes**. Instead, it returns a JSON error response.

### Example: Malformed JSON
```vbscript
Dim jsonParser, badJson, result

Set jsonParser = CreateObject("JSON.Parser")
badJson = "{invalid: json}"

Set result = jsonParser.ParseJson(badJson)

WScript.Echo result("error") ' Outputs: "Error: Invalid JSON format."

Set jsonParser = Nothing
Set result = Nothing
```

---

## 📌 Unregistering the Component
To remove the component, run:
```cmd
regsvr32 /u "C:\path\to\JSONParser.wsc"
```

---

## 🛠️ Troubleshooting
### ❓ Getting "ActiveX Can't Create Object" Error
- Ensure the WSC file is in a **trusted** location.
- Run `regsvr32` as **Administrator**.
- Verify with:
  ```cmd
  reg query "HKLM\Software\Classes\JSON.Parser"
  ```
  If it doesn't exist, re-register the component.

---

## 🌟 Why Use This Component?
- **No Third-Party Dependencies** – 100% Pure Classic ASP & VBScript.
- **Handles Complex JSON** – Nested objects, arrays, booleans, numbers, nulls.
- **Error-Proof** – Returns structured error messages.
- **Reusable Anywhere** – ASP, VBScript, WSH, Windows Services.

---

## 📜 License
This project is licensed under the **MIT License**.

---

## 🚀 Contributing
1. Fork the repository.
2. Improve the parser or add new features.
3. Submit a pull request!

💡 Feature Suggestions? Open an issue! 😊

---

## 👨‍💻 Maintainer
Developed and maintained by **Jose Gomez**.

---

## 💡 Next Steps
- ✅ Add **Pretty-Print JSON Formatting**
- ✅ Expand Support for **Advanced Data Structures**
- ✅ Optimize for **Massive JSON Files**

---

### 🎯 Ready to Use JSON in Classic ASP Like a Pro? Install & Go!
🔥 **Fast. Lightweight. Reliable.** 🚀
