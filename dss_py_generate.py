import os
import win32com.client
import importlib.util

class Config:
    pass

def read_settings(filepath):
    """ Ham doc file cau hinh settings.txt """
    cfg = Config()
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            exec(f.read(), {}, cfg.__dict__)
    return cfg

def extract_element_cards(opendss_link):
    """ Ham doc file Master va trich xuat the thong so phan tu (Name, Bus, P, Q, S, State, SOC) """
    d = win32com.client.Dispatch('OpenDSSEngine.DSS')
    d.Start(0)
    d.Text.Command = f"Compile '{opendss_link}'"
    
    cards = {'Generator': {}, 'PVSystem': {}, 'WindGen': {}, 'Storage': {}, 'Switch': {}, 
             'Load': {}, 'Capacitor': {}, 'Transformer': {}, 'Recloser': {}, 'Relay': {}}
    
    def scan_class(cls_name, category):
        d.ActiveCircuit.SetActiveClass(cls_name)
        names = d.ActiveCircuit.ActiveClass.AllNames
        if names and names[0].upper() != 'NONE':
            for name in names:
                d.ActiveCircuit.SetActiveElement(f"{cls_name}.{name}")
                bus = d.ActiveCircuit.ActiveCktElement.BusNames[0].split('.')[0]
                cards[category][name] = {
                    'Name': name,
                    'Bus': bus,
                    'P': 0.0,
                    'Q': 0.0,
                    'S': 0.0,
                    'State': 'ON',
                    'SOC': 100.0 if category == 'Storage' else None
                }

    scan_class('Generator', 'Generator')
    scan_class('WindGen', 'WindGen')
    scan_class('PVSystem', 'PVSystem')
    scan_class('Storage', 'Storage')
    scan_class('Capacitor', 'Capacitor')
    scan_class('Transformer', 'Transformer')
    scan_class('Recloser', 'Recloser')
    scan_class('Relay', 'Relay')
    
    # Chỉ quét SwtControl (bộ điều khiển công tắc) hoặc Line có thuộc tính switch=yes
    d.ActiveCircuit.SetActiveClass('Line')
    names = d.ActiveCircuit.ActiveClass.AllNames
    if names and names[0].upper() != 'NONE':
        for name in names:
            d.ActiveCircuit.SetActiveElement(f"Line.{name}")
            is_switch = d.ActiveCircuit.ActiveCktElement.Properties("switch").Val.lower() in ['y', 'yes', 'true']
            if is_switch:
                bus = d.ActiveCircuit.ActiveCktElement.BusNames[0].split('.')[0]
                cards['Switch'][name] = {'Name': name, 'Bus': bus, 'P': 0.0, 'Q': 0.0, 'S': 0.0, 'State': 'ON', 'SOC': None}

    return cards

def fetch_algorithm_data(algo_link, hour, element_category, element_name):
    """ Ham nhan du lieu tu thuat toan cho tung gio va tung phan tu """
    algo_path = algo_link
    if not os.path.isabs(algo_path):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        algo_path = os.path.join(current_dir, algo_link)
        
    if os.path.exists(algo_path):
        spec = importlib.util.spec_from_file_location("algo", algo_path)
        algo_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(algo_module)
        if hasattr(algo_module, 'get_data'):
            return algo_module.get_data(hour, element_category, element_name)
    return ""

def parse_power_string(raw_str):
    """ Ham phan tich chuoi du lieu cong suat thanh dictionary """
    parts = raw_str.replace('_', ' ').split()
    if len(parts) >= 3:
        return {'Name': parts[0], 'P': float(parts[1]), 'Q': float(parts[2])}
    return {}

def parse_switch_string(raw_str):
    """ Ham phan tich chuoi du lieu dong cat thanh dictionary """
    parts = raw_str.replace('_', ' ').split()
    if len(parts) >= 2:
        state = 'Open' if parts[1].lower() == 'noactive' else 'Close'
        return {'Name': parts[0], 'State': state}
    return {}

def parse_tap_string(raw_str):
    """ Ham phan tich chuoi du lieu tap cua may bien ap """
    parts = raw_str.replace('_', ' ').split()
    if len(parts) >= 2:
        return {'Name': parts[0], 'Tap': float(parts[1])}
    return {}

def translate_generator(name, data):
    """ Ham chuyen doi du lieu de nap vao may phat truyen thong """
    kw = data.get('P', 0)
    kvar = data.get('Q', 0)
    return f"edit Generator.{name} kW={kw} kvar={kvar}"

def translate_windgen(name, data):
    """ Ham chuyen doi du lieu de nap vao may phat dien gio """
    kw = data.get('P', 0)
    kvar = data.get('Q', 0)
    return f"edit WindGen.{name} kW={kw} kvar={kvar}"

def translate_pvsystem(name, data):
    """ Ham chuyen doi du lieu de nap vao pin mat troi """
    kw = data.get('P', 0)
    kvar = data.get('Q', 0)
    return f"edit PVSystem.{name} pmpp={kw} kvar={kvar}"

def translate_storage(name, data):
    """ Ham chuyen doi du lieu de nap vao pin luu tru """
    state = data.get('State', 'IDLE')
    soc = data.get('SOC', 100)
    kw = data.get('P', 0)
    return f"edit Storage.{name} state={state} kW={kw} %stored={soc}"

def translate_switch(name, data):
    """ Ham chuyen doi du lieu de nap vao thiet bi dong cat """
    state = data.get('State', 'Close')
    return f"{state} Line.{name} terminal=1"

def translate_load(name, data):
    """ Ham chuyen doi du lieu de nap vao phu tai """
    kw = data.get('P', 0)
    kvar = data.get('Q', 0)
    return f"edit Load.{name} kW={kw} kvar={kvar}"

def translate_capacitor(name, data):
    """ Ham chuyen doi du lieu de nap vao tu bu """
    kvar = data.get('Q', 0) # Tụ bù chủ yếu dùng Q
    return f"edit Capacitor.{name} kvar={kvar}"

def translate_transformer(name, data):
    """ Ham chuyen doi du lieu tap de nap vao may bien ap """
    tap = data.get('Tap', 1.0)
    return f"edit Transformer.{name} tap={tap}"

def translate_recloser(name, data):
    """ Ham chuyen doi du lieu de nap vao Recloser """
    state = data.get('State', 'Close')
    return f"{state} Recloser.{name}"

def translate_relay(name, data):
    """ Ham chuyen doi du lieu de nap vao Relay """
    state = data.get('State', 'Close')
    return f"{state} Relay.{name}"

def parse_stepsize_to_hours(stepsize_str):
    stepsize_str = str(stepsize_str).strip().lower()
    if stepsize_str.endswith('h'): return float(stepsize_str[:-1])
    elif stepsize_str.endswith('m'): return float(stepsize_str[:-1]) / 60.0
    elif stepsize_str.endswith('s'): return float(stepsize_str[:-1]) / 3600.0
    return 1.0

def build_snapshot_script():
    """ Ham xuat ket qua la mot file text dss thuc hien snapshot tung buoc """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    settings = read_settings(os.path.join(current_dir, 'settings.txt'))
    
    step_hours = parse_stepsize_to_hours(settings.STEPSIZE)
    
    cards = extract_element_cards(settings.OPENDSS_LINK)
    
    out_lines = []
    out_lines.append("clear")
    out_lines.append(f"Compile '{settings.OPENDSS_LINK}'")
    out_lines.append(f"Set mode=snap")
    out_lines.append("")
    
    for step in range(settings.NUMBERS):
        total_h = step * step_hours
        h = int(total_h)
        s = (total_h - h) * 3600
        out_lines.append(f"! === BUOC THOI GIAN {step + 1} ===")
        out_lines.append(f"set hour={h} sec={s}")
        
        for name in cards.get('Generator', {}):
            raw_str = fetch_algorithm_data(settings.ALGO_LINK, step, 'Generator', name)
            if raw_str: 
                data = parse_power_string(raw_str)
                out_lines.append(translate_generator(name, data))
                
        for name in cards.get('WindGen', {}):
            raw_str = fetch_algorithm_data(settings.ALGO_LINK, step, 'WindGen', name)
            if raw_str:
                data = parse_power_string(raw_str)
                out_lines.append(translate_windgen(name, data))
                
        for name in cards.get('PVSystem', {}):
            raw_str = fetch_algorithm_data(settings.ALGO_LINK, step, 'PVSystem', name)
            if raw_str:
                data = parse_power_string(raw_str)
                out_lines.append(translate_pvsystem(name, data))
                
        for name in cards.get('Storage', {}):
            raw_str = fetch_algorithm_data(settings.ALGO_LINK, step, 'Storage', name)
            if raw_str:
                data = parse_power_string(raw_str)
                out_lines.append(translate_storage(name, data))
                
        for name in cards.get('Capacitor', {}):
            raw_str = fetch_algorithm_data(settings.ALGO_LINK, step, 'Capacitor', name)
            if raw_str:
                data = parse_power_string(raw_str)
                out_lines.append(translate_capacitor(name, data))
                
        for name in cards.get('Transformer', {}):
            raw_str = fetch_algorithm_data(settings.ALGO_LINK, step, 'Transformer', name)
            if raw_str:
                data = parse_tap_string(raw_str)
                out_lines.append(translate_transformer(name, data))
                
        for name in cards.get('Switch', {}):
            raw_str = fetch_algorithm_data(settings.ALGO_LINK, step, 'Switch', name)
            if raw_str:
                data = parse_switch_string(raw_str)
                out_lines.append(translate_switch(name, data))
                
        for name in cards.get('Recloser', {}):
            raw_str = fetch_algorithm_data(settings.ALGO_LINK, step, 'Recloser', name)
            if raw_str:
                data = parse_switch_string(raw_str)
                out_lines.append(translate_recloser(name, data))
                
        for name in cards.get('Relay', {}):
            raw_str = fetch_algorithm_data(settings.ALGO_LINK, step, 'Relay', name)
            if raw_str:
                data = parse_switch_string(raw_str)
                out_lines.append(translate_relay(name, data))
            
        out_lines.append("Solve")
        out_lines.append("Sample")
        out_lines.append("")
        
    out_path = os.path.join(current_dir, 'output_dss_py.dss')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(out_lines))
