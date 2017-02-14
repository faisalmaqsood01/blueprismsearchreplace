# blueprismsearchreplace

' Input File/XML
StrFileName = "./test2.xml"
 
' Output XML
result = ""
 
' Text data items and text collection items
stagetypes = "|Data|Collection|"
dataTypes = "|text|collection|"
 
' Search For (original initial value text) and Replace With (revised initial value text)
searchFor = "\\ahwfilinmsc001\departments$\"
replaceWith = "\\inmsc.ds.sjhs.com\departments$\"
 
Set ObjFso = CreateObject("Scripting.FileSystemObject")
Set ObjFile = ObjFso.OpenTextFile(StrFileName)
MyVar = ObjFile.ReadAll
ObjFile.Close
 
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.setProperty "SelectionLanguage", "XPath"
xmlDoc.async="false"
xmlDoc.load( StrFileName )
 
strXPath = "/process/stage"
Set processes = xmlDoc.documentElement.childNodes
set xmlDoc2 = CreateObject("Microsoft.XMLDOM")
Set stages = xmlDoc.selectNodes("/process/stage")
 
 
For Each stage in stages
 
  xmlDoc2.loadXML( stage.xml )
 
  ' Comment: Type of Blue Prism Stage
  If InStr( stagetypes, ( "|" & stage.getAttribute("type") & "|" )) > 0  Then
 
 
    If xmlDoc2.selectSingleNode("stage/datatype").Text = "text"  Then
      If Not xmlDoc2.selectSingleNode("stage/initialvalue") Is Nothing Then
 
        currentValue = xmlDoc2.selectSingleNode("stage/initialvalue").Text
        revisedValue = replaceTextValues( currentValue, searchFor, replaceWith )
        xmlDoc2.selectSingleNode("stage/initialvalue").Text = revisedValue
 
      End If
    End If
 
    If Not xmlDoc2.selectSingleNode("stage/initialvalue/row") Is Nothing Then
 
      Set fields = xmlDoc2.selectNodes( "stage/initialvalue/row/field" )
                            
      For Each field in fields
        If InStr( dataTypes, ( "|" & field.getAttribute("type") & "|" )) > 0  Then
 
          currentValue = field.getAttribute( "value" )
          replaceTextValues = Replace( string, searchFor, replaceWith, 1, -1, vbTextCompare )
          field.setAttribute "value", revisedValue
 
        End If
      Next ' (field)
 
  End If
 
 
End If
 
result = result & xmlDoc2.xml
 
Next ' (stage)
 
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
 
Set objTextFile = objFSO.OpenTextFile("./replace-results.xml", ForWriting, True)
 
  objTextFile.WriteLine(result)
  objTextFile.Close
  Set ObjFso = Nothing
 
function replaceTextValues( string, searchFor, replaceWith )
 
    replaceTextValues = Replace( string, searchFor, replaceWith )
 
end function
