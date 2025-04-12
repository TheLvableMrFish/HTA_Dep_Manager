function loadFileList(){
    var directory = document.getElementById("directoryInput").value
    var filterText = document.getElementById("filterInput").value.toLowerCase()
    var fileList = document.getElementById("fileList")

    fileList.options.length = 0

    try{
        var fileSysObj = new ActiveXObject("Scripting.FileSystemObject")
        var folder = fileSysObj.GetFolder(directory)
        var files = new Enumerator(folder.Files)

        for(; !files.atEnd(); files.moveNext()){
            var file = files.item()
            if(filterText === "" || file.Name.toLowerCase().indexOf(filterText) !== -1){
                var option = document.createElement("option")
                option.text = file.Name
                option.value = file.Path
                fileList.add(option)
            }
        }
    } catch(e){
        alert("Error loading file list: " + e.message);
    }
}

// When a file is selected from the list, update the file path input
function selectFileFromList(){
    var fileList = document.getElementById("fileList")
    if(fileList.selectedIndex >= 0){
        var filePath = fileList.options[fileList.selectedIndex].value
        document.getElementById("filePathInput").value = filePath
    }
}

// Loads the selected file into the editor, creates txt file if none of the name exist
function loadFile(){
    var filePath = document.getElementById("filePathInput").value
    
    if(filePath === ""){
        alert("Please enter or select a file path.")
        return
    }

    try{
        var fileSysObj = new ActiveXObject("Scripting.FileSystemObject")
        if(!fileSysObj.FileExists(filePath)){
            var newFile = fileSysObj.CreateTextFile(filePath, true)
            newFile.WriteLine("This is your new file! Edit as needed.")
            newFile.Close()
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
    var filePath = document.getElementById("filePathInput").value
    if(filePath === ""){
        alert("Please enter or select a file path.")
        return
    }
    try{
        var fileSysObj = new ActiveXObject("Scripting.FileSystemObject")
        var file = fileSysObj.CreateTextFile(filePath, true) // overwrites file
        file.Write(document.getElementById("editor").value)
        file.Close()
        alert("File saved successfully!")
    } catch(e){
        alert("Error saving file: " + e.message)
    }
}