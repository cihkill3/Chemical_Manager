import re

def parse_hazard(hazard_str):
    """
    Parses a combined hazard string (e.g. "[Danger] H302, H315...") into:
    - 신호어 (Signal Word) with emoji
    - 주요위험 (Main Hazard) with emoji
    - 상세 위험분류 (Detailed Hazard)
    """
    if not hazard_str or hazard_str == "정보 없음":
        return "정보 없음", "정보 없음", "정보 없음"
        
    signal_word = "정보 없음"
    main_hazard = "정보 없음"
    detailed_hazard = hazard_str
    
    # 1. 신호어 추출
    signal_match = re.search(r'\[(.*?)\]', hazard_str)
    if signal_match:
        signal_raw = signal_match.group(1).lower()
        if 'danger' in signal_raw or '위험' in signal_raw:
            signal_word = "● 위험"
        elif 'warning' in signal_raw or '경고' in signal_raw:
            signal_word = "▲ 경고"
        elif 'none' in signal_raw or '해당없음' in signal_raw or '없음' in signal_raw:
            signal_word = "■ 안전"
        else:
            signal_word = f"◆ {signal_match.group(1)}"
            
        # 상세 위험에서 신호어 부분 제거
        detailed_hazard = detailed_hazard.replace(f"[{signal_match.group(1)}]", "").strip()
        
    # 만약 "GHS 기준에 의거하여 유해화학물질로 분류되지 않음" 같은 문구가 있다면 안전으로 처리
    if "분류되지 않음" in detailed_hazard or "not classified" in detailed_hazard.lower():
        if signal_word == "정보 없음":
            signal_word = "■ 안전"
        if main_hazard == "정보 없음":
            main_hazard = "■ 해당없음"
            
    # 2. 주요위험 추출 (간단한 키워드 매칭)
    hazard_lower = detailed_hazard.lower()
    hazards = []
    
    if 'toxic' in hazard_lower or '독성' in hazard_lower or 'fatal' in hazard_lower or '치명적' in hazard_lower:
        hazards.append("● 독성")
    if 'corrosive' in hazard_lower or '부식' in hazard_lower or 'burns' in hazard_lower or '화상' in hazard_lower or '눈 손상' in hazard_lower or 'damage' in hazard_lower:
        hazards.append("▲ 부식성/눈손상")
    if 'flammable' in hazard_lower or '인화성' in hazard_lower or 'fire' in hazard_lower or '화재' in hazard_lower:
        hazards.append("★ 인화성")
    if 'irrit' in hazard_lower or '자극' in hazard_lower or 'sensit' in hazard_lower or '과민성' in hazard_lower or '졸음' in hazard_lower:
        hazards.append("◆ 자극성/과민성")
    if 'environmental' in hazard_lower or 'aquatic' in hazard_lower or '수생' in hazard_lower or '환경' in hazard_lower:
        hazards.append("■ 환경유해")
    if 'health' in hazard_lower or '건강' in hazard_lower or 'cancer' in hazard_lower or '발암' in hazard_lower or '장기' in hazard_lower or '생식' in hazard_lower or '유전' in hazard_lower:
        hazards.append("▼ 건강유해")
        
    if hazards:
        main_hazard = ", ".join(hazards)
        
    return signal_word, main_hazard, detailed_hazard
