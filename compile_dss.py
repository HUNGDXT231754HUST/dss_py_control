import os
import win32com.client
import main as snapshot_bridge
import shutil
import glob

def create_monitors(d):
    """ Tao Monitor """
    monitor_lines = []
    classes_to_monitor = ['Line', 'Transformer', 'Generator', 'PVSystem', 'Storage', 'Capacitor', 'Load']
    
    for cls in classes_to_monitor:
        d.ActiveCircuit.SetActiveClass(cls)
        names = d.ActiveCircuit.ActiveClass.AllNames
        if names and names[0].upper() != 'NONE':
            for name in names:
                monitor_lines.append(f"New Monitor.{cls}_{name} element={cls}.{name} terminal=1 mode=0")
                
    return monitor_lines

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 1. Chay snapshot_bridge
    print("Dang chay file snapshot...")
    snapshot_bridge.build_snapshot_script()
    
    out_file = os.path.join(current_dir, 'output_snapshots.dss')
    if not os.path.exists(out_file):
        return
        
    # 2. Khai bao Monitor
    settings = snapshot_bridge.read_settings(os.path.join(current_dir, 'settings.txt'))
    
    d = win32com.client.Dispatch('OpenDSSEngine.DSS')
    d.Start(0)
    d.Text.Command = f"Compile '{settings.OPENDSS_LINK}'"
    
    monitor_cmds = create_monitors(d)
    
    # Chen lenh Monitor
    with open(out_file, 'r', encoding='utf-8') as f:
        content = f.read()
        
    monitor_block = "\n! --- Khai bao Monitors ---\n" + "\n".join(monitor_cmds) + "\n"
    content = content.replace("Set mode=snap", monitor_block + "Set mode=snap")
    
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(content)
        
    # 3. Chay file DSS
    print(f"Dang giai tich...")
    d.Text.Command = f"Redirect '{out_file}'"
    
    # 4. Xuat CSV
    d.Text.Command = "Export monitors all"
    
    csv_dir = os.path.join(current_dir, 'csv')
    dsviz_dir = os.path.join(current_dir, 'dsviz')
    
    os.makedirs(csv_dir, exist_ok=True)
    os.makedirs(dsviz_dir, exist_ok=True)
    
    # Tim file CSV
    data_path = d.Text.Command = "get datapath"
    if not data_path:
        data_path = os.path.dirname(settings.OPENDSS_LINK)
        
    csv_files = glob.glob(os.path.join(data_path, "*Mon_*.csv"))
    for f in csv_files:
        try:
            shutil.move(f, os.path.join(csv_dir, os.path.basename(f)))
        except:
            pass
            
    print("Hoan tat!")

if __name__ == '__main__':
    main()
