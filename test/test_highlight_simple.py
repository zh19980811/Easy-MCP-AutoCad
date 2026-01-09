import win32com.client
import pythoncom

def test_highlight_simple(handle, color=3):
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        
        print(f"Attempting to get entity with handle {handle}...")
        entity = doc.HandleToObject(handle)
        
        old_color = entity.Color
        entity.Color = color
        
        print(f"Success! Highlighted entity {handle}. Color changed from {old_color} to {color}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    # 80 is the handle of the new circle
    test_highlight_simple("80", 1)  # Color 1 is Red
