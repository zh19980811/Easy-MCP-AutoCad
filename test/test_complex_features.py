import win32com.client
import pythoncom
import math

def test_features():
    print("Connecting to AutoCAD...")
    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    model_space = doc.ModelSpace
    
    # 1. Test Layer
    print("\n--- Testing Create Layer ---")
    layer_name = "MCP_Test_Layer"
    try:
        layer = doc.Layers.Item(layer_name)
    except:
        layer = doc.Layers.Add(layer_name)
    layer.Color = 1 # Red
    doc.ActiveLayer = layer
    print(f"Layer '{layer_name}' created/set active.")

    # 2. Test Polyline/Rectangle
    print("\n--- Testing Draw Rectangle (Polyline) ---")
    # Rectangle (0,0) to (20,10)
    points = [0, 0, 20, 0, 20, 10, 0, 10, 0, 0] # Closed loop
    # Variant conversion
    pt_array = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, points)
    rect = model_space.AddLightWeightPolyline(pt_array)
    rect.Closed = True
    print(f"Rectangle created. Handle: {rect.Handle}")
    
    # 3. Test Text
    print("\n--- Testing Draw Text ---")
    insert_pnt = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [5, 5, 0])
    text = model_space.AddText("MCP Text", insert_pnt, 2.5)
    print(f"Text created. Handle: {text.Handle}")
    
    # 4. Test Move
    print("\n--- Testing Move Entity ---")
    # Move rectangle from (0,0) to (50,50)
    p1 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [0, 0, 0])
    p2 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [50, 50, 0])
    rect.Move(p1, p2)
    print("Rectangle moved.")
    
    # 5. Test Rotate
    print("\n--- Testing Rotate Entity ---")
    # Rotate rectangle 45 degrees around (50,50)
    p_base = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [50, 50, 0])
    angle_rad = 45 * math.pi / 180
    rect.Rotate(p_base, angle_rad)
    print("Rectangle rotated 45 degrees.")
    
    # 6. Test Copy
    print("\n--- Testing Copy Entity ---")
    # Copy text from (5,5) -> (55, 55) (since we moved rect, let's copy text to follow it roughly)
    # Text is at (5,5). Let's copy it to (60, 60).
    text_copy = text.Copy()
    p_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [0, 0, 0]) # Relative move? No, Move takes points.
    # Copy() makes it in place. So we move it from its current position.
    # Current text pos is (5,5). We want it at (60,60). Vector is (55,55).
    # So move from (0,0) to (55,55) is one way.
    p_to = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [55, 55, 0])
    text_copy.Move(p_start, p_to)
    print(f"Text copied. New Handle: {text_copy.Handle}")
    
    print("\nAll tests completed.")

if __name__ == "__main__":
    test_features()
