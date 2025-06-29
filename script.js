// Initialize shell and get env var for the dep folder
var shell = new ActiveXObject("WScript.Shell")
var env = shell.Environment("USER")
var Dep_Folder_Path = env("Dep_Folder_Path")

// Get curr directory wher the HTA file is located
var fileSysObj = new ActiveXObject("Scripting.FileSystemObject")
var currentDir = shell.CurrentDirectory
var txtPath = currentDir + "\\depList.txt"

// Dark mode
var darkMode = true

// Load file list from the depList.txt from the curr directory 
window.onload = function (){
    if(!Dep_Folder_Path){
        alert("Dep_Folder_Path environment variable is not set")
    }
    GenerateDepList()
}

// Generates list to select from
function GenerateDepList (){
    try{
        if(fileSysObj.FileExists(txtPath)){
            // For Reading
            var file = fileSysObj.OpenTextFile(txtPath, 1)
            var fileList = document.getElementById("fileList")
            fileList.options.length = 0

            // Read each line for each file name saved in list
            while(!file.AtEndOfStream){
                var line = file.ReadLine()
                if(line.length > 0){
                    var option = document.createElement("option")
                    option.text = line
                    option.value = Dep_Folder_Path + "\\" + line
                    fileList.add(option)
                }
            }
            file.Close()
        } else {
            alert("depList.txt not found next to .hta file.")
        }
    } catch(e){
        alert("Error loading depList.txt " + e.message)
    }
}

// Update file path input when selecting from dropdown
function selectFileFromList(){
    var fileList = document.getElementById("fileList")
    if(fileList.selectedIndex >= 0){
        var filePath = fileList.options[fileList.selectedIndex].text
        var fileNameOnly = fileSysObj.GetBaseName(filePath)
        document.getElementById("filePathInput").value =  fileNameOnly 
    }
}

// Load the list of files from the actual dependency folder
function loadFileList(){
    var directory = Dep_Folder_Path
    var filterText = document.getElementById("filterInput").value.toLowerCase()
    var fileList = document.getElementById("fileList")

    fileList.options.length = 0

    try{
        var folder = fileSysObj.GetFolder(directory)
        var files = new Enumerator(folder.Files)

        // Loop through the dep files in the dep folder
        for(; !files.atEnd(); files.moveNext()){
            var file = files.item()
            if(filterText === "" || file.Name.toLowerCase().indexOf(filterText) !== -1){
                var option = document.createElement("option")
                option.text = fileSysObj.GetBaseName(file.Name)
                option.value = file.Path
                fileList.add(option)
            }
        }
    } catch(e){
        alert("Error loading file list: " + e.message);
    }
}



// Loads the selected file into the editor, creates txt file if none of the name exist
// File doesn't exist create a new file wit hthe name
function loadFile(){
    var fileName = document.getElementById("filePathInput").value
    var filePath = Dep_Folder_Path + "\\" + fileName + ".csv"
    
    if(fileName === ""){
        alert("Please enter or select a file name.")
        return
    }

    try{
        if(!fileSysObj.FileExists(filePath)){
            // var newFile = fileSysObj.CreateTextFile(filePath, true)
            // newFile.WriteLine("")
            // newFile.Close()
        }
        var file = fileSysObj.OpenTextFile(filePath, 1)
        var content = ""

        if(!file.AtEndOfStream){
            // for empty files
           var content = file.ReadAll()
        }
        file.Close()
        document.getElementById("editor").value = content

    } catch(e){
        alert("errors loading file: " + e.message)
    }
}

function saveFile(){
    var fileName = document.getElementById("filePathInput").value
    
    if(fileName === ""){
        alert("Please enter or select a file Name.")
        return
    }

    if(fileName.substring(fileName.length - 4) !== ".csv"){
        fileName = fileName + ".csv"
    } 

    var filePath = Dep_Folder_Path + "\\" + fileName

    try{
        var fileSysObj = new ActiveXObject("Scripting.FileSystemObject")
        var file = fileSysObj.CreateTextFile(filePath, true) // overwrites file
        file.Write(document.getElementById("editor").value)
        file.Close()

        addFileToList(fileName)
        
        alert("File saved successfully!")
    } catch(e){
        alert("Error saving file: " + e.message)
    }
}

function addFileToList(fileName){
    try{
        var exists = false

        if(fileSysObj.FileExists(txtPath)){
            var readFile = fileSysObj.OpenTextFile(txtPath,1)
            while(!readFile.AtEndOfStream){
                var line = readFile.Readline()
                var noSpaceLine = line.replace(/^\s+|\s$/g, "")
                if(noSpaceLine.toLowerCase() === fileName.toLowerCase()){
                    exists = true
                    break
                }
            }
            readFile.Close()
        }

        if(!exists){

            if(fileSysObj.FileExists(txtPath)){
                // 8 is Append mode var
                var file = fileSysObj.OpenTextFile(txtPath, 8)
                file.WriteLine(fileName)
                file.Close()
            } else {
                var file = fileSysObj.CreateTextFile(txtPath, true)
                file.WriteLine(fileName)
                file.Close()
            }
            GenerateDepList()
        }        

    } catch(e){
        alert("Error adding file to depList.txt" + e.message)
    }
}

function refreshDepList(){
    try{
        var folder = fileSysObj.GetFolder(Dep_Folder_Path)
        var files = new Enumerator(folder.Files)
        var listFile = fileSysObj.CreateTextFile(txtPath, true)

        for(; !files.atEnd(); files.moveNext()){
            var file = files.item()
            listFile.WriteLine(file.Name)
        }

        listFile.Close()
        alert("depList.txt refreshed!")
        GenerateDepList()
    } catch(e){
        alert("Error updating the depList.txt" + e.message)
    }
}



function generateNewFile(){
    var fileName = document.getElementById("filePathInput").value

    if(fileName === ""){
        alert("Please enter or select a file name.")
        return
    }
    
    if(fileName.substring(fileName.length - 4) !== ".csv"){
        fileName += ".csv"
    }

    var filePath = Dep_Folder_Path + "\\" + fileName

    try{
        var fileSysObj = new ActiveXObject("Scripting.FileSystemObject")
        
        if(!fileSysObj.FileExists(filePath)){
            var newFile = fileSysObj.CreateTextFile(filePath, true)

            if(document.getElementById("editor").value === ""){
                newFile.WriteLine("Dep info here")
            } else {
                newFile.Write(document.getElementById("editor").value)
            }

            newFile.Close()
            addFileToList(fileName)
            alert("File created successfully!")

        } else {
            alert("File alread exists")
        }
        
    } catch(e){
        alert("Error saving file: " + e.message)
    }
}

function toggleMode(mode){
    var selectBtn = document.getElementById('selectModeBtn')
    var editBtn = document.getElementById('editModeBtn')

    var editMode = document.getElementById('editMode')
    var selectMode = document.getElementById('selectMode')

    if(mode === 'select'){
        selectBtn.className = 'dep-toggle toggle-on'
        editBtn.className = 'dep-toggle toggle-off'
        selectMode.className = 'select-mode'
        editMode.className = 'edit-mode hidden'
    } else {
        selectBtn.className = 'dep-toggle toggle-off'
        editBtn.className = 'dep-toggle toggle-on'
        editMode.className = 'edit-mode'
        selectMode.className = 'select-mode hidden'
    }
}

function toggleDarkMode(){
    darkMode = !darkMode
    document.body.style.backgroundColor = darkMode ? '#1E1E1E' : 'white'
    document.body.style.color = darkMode ? '#CE9178' : 'black'
    document.documentElement.style.backgroundColor = darkMode ? '#1E1E1E' : 'white'
    document.documentElement.style.color = darkMode ? '#CE9178' : 'black'
    
    var darkBtn = document.getElementById('dark-btn')
    var lightBtn = document.getElementById('light-btn')
    var textBox = document.getElementById('editor')
    var fileList = document.getElementById('fileList')

    darkBtn.className = darkMode ? 'dark btn' : 'hidden'
    lightBtn.className = !darkMode ? 'light btn' : 'hidden'
    textBox.className = darkMode ? 'content-box content-box-dark' : 'content-box content-box-light'
    fileList.className = darkMode ? 'dep-items dep-items-dark' : 'dep-items dep-items-light'
}

function createTxtFile(){
    var option1 = document.getElementById('option1').value
    var option2 = document.getElementById('option2').value
    var option3 = document.getElementById('option3').value
    var option4 = document.getElementById('option4').value
    var option5 = document.getElementById('option5').value
    var option6 = document.getElementById('option6').value
    var option7 = document.getElementById('option7').value
    var option8 = document.getElementById('option7').value

    var file = 'templates first line'
    var firstLine = "text"
    var optionsList1 = [option1, option2, option3, option4, option5, option6, option7, option8]
    var optionsList2 = []
    var hasValues = false

    for(var i = 0; i < optionsList1.length; i++){
        optionsList1[i] == "" ? null : hasValues = true
    }
    var fullText
}