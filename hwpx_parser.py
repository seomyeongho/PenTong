import zipfile
import xml.etree.ElementTree as ET

class HwpXParser:
    def extract_text(self, hwpx_path):
        """HWPX 파일 내부의 xml을 파싱하여 텍스트와 비율/병합/투명선 및 중첩표가 완벽히 적용된 직사각형 표를 추출합니다."""
        ns_hp = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
        result_lines = []
        
        try:
            with zipfile.ZipFile(hwpx_path, 'r') as zf:
                # 1. header.xml 테두리 파싱
                border_fills = {}
                try:
                    header_root = ET.fromstring(zf.read('Contents/header.xml'))
                    for bf in header_root.iter():
                        if bf.tag.endswith('borderFill'):
                            bf_id = bf.get('id')
                            borders = {}
                            for child in bf:
                                tag_name = child.tag.split('}')[-1]
                                if tag_name in ['leftBorder', 'rightBorder', 'topBorder', 'bottomBorder']:
                                    b_type = child.get('type', 'solid')
                                    side = tag_name.replace('Border', '')
                                    borders[side] = 'none' if b_type.lower() == 'none' else '1px solid black'
                                elif tag_name == 'fillBrush':
                                    for brush in child:
                                        b_tag = brush.tag.split('}')[-1]
                                        if b_tag == 'solidBrush':
                                            borders['bg'] = brush.get('color', 'transparent')
                                        elif b_tag == 'winBrush':
                                            borders['bg'] = brush.get('faceColor', 'transparent')
                            border_fills[bf_id] = borders
                except Exception:
                    pass

                sections = sorted([f for f in zf.namelist() if f.startswith('Contents/section') and f.endswith('.xml')])
                
                # 💡 모든 요소를 재귀적으로 탐색하는 메인 파서 함수
                def parse_element(elem):
                    chunks = []
                    
                    if elem.tag == f"{{{ns_hp}}}tbl":
                        trs = elem.findall(f"{{{ns_hp}}}tr")
                        if not trs: return []
                        
                        # 전체 구조 시뮬레이션: 표의 '절대 최대 너비' 계산
                        max_c = 0
                        temp_occupied = set()
                        for r, tr in enumerate(trs):
                            c = 0
                            for tc in tr.findall(f"{{{ns_hp}}}tc"):
                                while (r, c) in temp_occupied: c += 1
                                colspan = int(tc.get("colSpan", "1"))
                                rowspan = int(tc.get("rowSpan", "1"))
                                for rr in range(rowspan):
                                    for cc in range(colspan):
                                        temp_occupied.add((r + rr, c + cc))
                                c += colspan
                            if c > max_c: max_c = c
                        num_cols = max_c

                        # 열별 너비 추출 및 빈칸 비율 배분
                        col_widths = [0] * num_cols
                        grid_for_w = set()
                        
                        for r, tr in enumerate(trs):
                            c = 0
                            for tc in tr.findall(f"{{{ns_hp}}}tc"):
                                while (r, c) in grid_for_w: c += 1
                                if c >= num_cols: break
                                colspan = int(tc.get("colSpan", "1"))
                                rowspan = int(tc.get("rowSpan", "1"))
                                
                                w = 0
                                tc_pr = tc.find(f".//{{{ns_hp}}}tcPr")
                                if tc_pr is not None:
                                    sz = tc_pr.find(f".//{{{ns_hp}}}cellSz")
                                    if sz is not None: w = int(sz.get("width", "0"))
                                        
                                if colspan == 1 and w > col_widths[c]:
                                    col_widths[c] = w
                                    
                                for rr in range(rowspan):
                                    for cc in range(colspan):
                                        grid_for_w.add((r + rr, c + cc))
                                c += colspan
                                
                        grid_for_w.clear()
                        for r, tr in enumerate(trs):
                            c = 0
                            for tc in tr.findall(f"{{{ns_hp}}}tc"):
                                while (r, c) in grid_for_w: c += 1
                                if c >= num_cols: break
                                colspan = int(tc.get("colSpan", "1"))
                                rowspan = int(tc.get("rowSpan", "1"))
                                
                                if colspan > 1:
                                    w = 0
                                    tc_pr = tc.find(f".//{{{ns_hp}}}tcPr")
                                    if tc_pr is not None:
                                        sz = tc_pr.find(f".//{{{ns_hp}}}cellSz")
                                        if sz is not None: w = int(sz.get("width", "0"))
                                    
                                    zeros = [c + i for i in range(colspan) if c + i < num_cols and col_widths[c + i] == 0]
                                    if zeros and w > 0:
                                        per = w / len(zeros)
                                        for z in zeros: col_widths[z] = per
                                for rr in range(rowspan):
                                    for cc in range(colspan):
                                        grid_for_w.add((r + rr, c + cc))
                                c += colspan
                                
                        # HTML 브라우저용 뼈대 비율(<colgroup>) 생성
                        total_w = sum(col_widths)
                        if total_w == 0 and num_cols > 0:
                            col_widths = [1] * num_cols
                            total_w = num_cols
                            
                        colgroup_html = "<colgroup>"
                        for w in col_widths:
                            pct = (w / total_w) * 100 if total_w else 0
                            colgroup_html += f"<col style='width: {pct:.2f}%;'>"
                        colgroup_html += "</colgroup>"
                        
                        html_table = ["<table style='width: 100%; table-layout: fixed; border-collapse: collapse; empty-cells: show; margin: 2px 0; font-size: 7pt; line-height: 1.3;'>"]
                        html_table.append(colgroup_html)
                        
                        # 표 렌더링 및 가짜 셀 스킵
                        occupied = set()
                        for r, tr in enumerate(trs):
                            html_table.append("<tr>")
                            c = 0
                            for tc in tr.findall(f"{{{ns_hp}}}tc"):
                                while (r, c) in occupied: c += 1
                                if c >= num_cols: break
                                
                                colspan = int(tc.get("colSpan", "1"))
                                rowspan = int(tc.get("rowSpan", "1"))
                                
                                border_css = "border: 1px solid black;"
                                tc_pr = tc.find(f".//{{{ns_hp}}}tcPr")
                                if tc_pr is not None:
                                    bf_id = tc_pr.get("borderFillIDRef")
                                    if bf_id and bf_id in border_fills:
                                        b = border_fills[bf_id]
                                        border_css = f"border-left: {b.get('left', '1px solid black')}; border-right: {b.get('right', '1px solid black')}; border-top: {b.get('top', '1px solid black')}; border-bottom: {b.get('bottom', '1px solid black')}; background-color: {b.get('bg', 'transparent')};"
                                
                                # 💡 [핵심 수정] 셀(tc) 내부를 무작정 글자만 긁어오지 않고, 재귀 파서(parse_element)에 통째로 넘깁니다.
                                cell_contents = []
                                sublist = tc.find(f"{{{ns_hp}}}subList")
                                target_container = sublist if sublist is not None else tc
                                
                                for p_elem in target_container.findall(f"{{{ns_hp}}}p"):
                                    p_parts = parse_element(p_elem) # 중첩 표가 있으면 이 안에서 표 HTML을 리턴함
                                    if p_parts:
                                        for part in p_parts:
                                            cell_contents.append(part)
                                        cell_contents.append("<br>") # 문단 구분을 위한 줄바꿈
                                        
                                # 마지막에 추가된 불필요한 줄바꿈 제거
                                if cell_contents and cell_contents[-1] == "<br>":
                                    cell_contents.pop()
                                    
                                final_html = "".join(cell_contents)
                                # 표와 표 사이, 또는 표 위아래에 생기는 불필요한 공백/줄바꿈 정리
                                final_html = final_html.replace("<br><table", "<table").replace("</table><br>", "</table>")
                                content = final_html if final_html.strip() else "&nbsp;"
                                
                                html_table.append(f"<td colspan='{colspan}' rowspan='{rowspan}' style='{border_css} padding: 1px 2px; text-align: center; vertical-align: middle; word-break: break-all; overflow-wrap: break-word;'>{content}</td>")
                                
                                for rr in range(rowspan):
                                    for cc in range(colspan):
                                        occupied.add((r + rr, c + cc))
                                c += colspan
                                
                            # 직사각형 이빨 빠짐 방지 패딩
                            while c < num_cols:
                                if (r, c) not in occupied:
                                    html_table.append("<td style='border: 1px solid black; padding: 1px 2px;'>&nbsp;</td>")
                                    occupied.add((r, c))
                                c += 1
                                
                            html_table.append("</tr>")
                        html_table.append("</table>")
                        return ["".join(html_table)]
                        
                    elif elem.tag == f"{{{ns_hp}}}p":
                        p_chunks = []
                        current_text = ""
                        for child in elem: 
                            if child.tag == f"{{{ns_hp}}}run":
                                for sub in child:
                                    if sub.tag == f"{{{ns_hp}}}t" and sub.text:
                                        current_text += sub.text
                                    elif sub.tag == f"{{{ns_hp}}}tbl":
                                        if current_text.strip(): 
                                            p_chunks.append(current_text.strip())
                                            current_text = ""
                                        p_chunks.extend(parse_element(sub))
                                    elif sub.tag == f"{{{ns_hp}}}ctrl":
                                        for c in sub:
                                            if c.tag == f"{{{ns_hp}}}tbl":
                                                if current_text.strip(): 
                                                    p_chunks.append(current_text.strip())
                                                    current_text = ""
                                                p_chunks.extend(parse_element(c))
                        if current_text.strip():
                            p_chunks.append(current_text.strip())
                        return p_chunks
                        
                    return chunks

                for sec in sections:
                    root = ET.fromstring(zf.read(sec))
                    for elem in root:
                        result_lines.extend(parse_element(elem))
                        
        except Exception as e:
            return f"HWPX 파싱 오류: {str(e)}"
            
        return "\n\n".join(result_lines)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    