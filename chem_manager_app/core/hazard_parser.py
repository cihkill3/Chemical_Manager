import re

def parse_hazard(detailed_hazard, input_signal_word=""):
    """
    Parses a combined detailed hazard string (e.g. "H302: ... / H314: ...") into:
    - Signal Word (with emoji symbol) in English
    - Key Hazards (with emoji symbol) in English
    - Detailed Hazard Classification
    """
    if not detailed_hazard or detailed_hazard in ["정보 없음", "N/A", "None", ""]:
        return input_signal_word if input_signal_word not in ["정보 없음", "N/A", "None"] else "", "", ""
        
    signal_word = input_signal_word if input_signal_word not in ["정보 없음", "N/A", "None"] else ""
    key_hazards = ""
    
    # 1. Signal Word logic
    if not signal_word:
        hazard_lower = detailed_hazard.lower()
        if 'danger' in hazard_lower or '위험' in hazard_lower or 'fatal' in hazard_lower or '치명' in hazard_lower or 'severe' in hazard_lower:
            signal_word = "● Danger"
        elif 'warning' in hazard_lower or '경고' in hazard_lower or 'harmful' in hazard_lower or 'irritat' in hazard_lower:
            signal_word = "▲ Warning"
        elif '분류되지 않음' in hazard_lower or 'not classified' in hazard_lower:
            signal_word = "■ Safe"
        else:
            signal_word = ""
    else:
        sw_lower = signal_word.lower()
        if 'danger' in sw_lower or '위험' in sw_lower:
            signal_word = "● Danger"
        elif 'warning' in sw_lower or '경고' in sw_lower:
            signal_word = "▲ Warning"
        elif 'safe' in sw_lower or '안전' in sw_lower or 'none' in sw_lower:
            signal_word = "■ Safe"
            
    if "분류되지 않음" in detailed_hazard or "not classified" in detailed_hazard.lower():
        if not signal_word:
            signal_word = "■ Safe"
            
    # 2. Key Hazards extraction (Keyword matching)
    hazard_lower = detailed_hazard.lower()
    hazards = []
    
    if 'toxic' in hazard_lower or '독성' in hazard_lower or 'fatal' in hazard_lower or '치명' in hazard_lower:
        hazards.append("● Toxic")
    if 'corrosive' in hazard_lower or '부식' in hazard_lower or 'burns' in hazard_lower or '화상' in hazard_lower or 'damage' in hazard_lower or '손상' in hazard_lower:
        hazards.append("▲ Corrosive / Eye Damage")
    if 'flammable' in hazard_lower or '인화성' in hazard_lower or 'fire' in hazard_lower or '화재' in hazard_lower:
        hazards.append("★ Flammable")
    if 'irrit' in hazard_lower or '자극' in hazard_lower or 'sensit' in hazard_lower or '과민성' in hazard_lower or '졸음' in hazard_lower:
        hazards.append("◆ Irritant / Sensitizer")
    if 'environmental' in hazard_lower or 'aquatic' in hazard_lower or '수생' in hazard_lower or '환경' in hazard_lower:
        hazards.append("■ Environmental Hazard")
    if 'health' in hazard_lower or '건강' in hazard_lower or 'cancer' in hazard_lower or '발암' in hazard_lower or '장기' in hazard_lower or '생식' in hazard_lower or '유전' in hazard_lower:
        hazards.append("▼ Health Hazard")
        
    if hazards:
        key_hazards = ", ".join(hazards)
        
    return signal_word, key_hazards, detailed_hazard
