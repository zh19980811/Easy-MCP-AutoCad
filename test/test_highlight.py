import win32com.client
import pythoncom

def test_highlight(handle, color=1):
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        
        print("Cleaning up potentially existing selection set...")
        # Cleanup
        try:
            doc.SelectionSets.Item("TempSS").Delete()
            print("Deleted existing TempSS")
        except:
            print("TempSS did not exist")
            pass
            
        print("Creating selection set...")
        selection = doc.SelectionSets.Add("TempSS")
        
        print("Selecting entity...")
        # Select
        filter_type = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
        filter_data = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["HANDLE", handle])
        
        # Mode 5 = acSelectionSetAll
        selection.Select(5, 0, 0, filter_type, filter_data) 
        
        if selection.Count == 0:
            print(f"No entity found with handle {handle}")
            selection.Delete()
            return
            
        print(f"Found {selection.Count} entities")
        entity = selection.Item(0)
        old_color = entity.Color
        entity.Color = color
        print(f"Success! Highlighted entity {handle}. Color changed from {old_color} to {color}")
        
        selection.Delete()
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    # 7F is the handle of the line we drew earlier
    test_highlight("7F", 1) 
