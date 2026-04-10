const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, LevelFormat } = require('docx');
const fs = require('fs');

const ACCENT = '6C5CE7';
const ACCENT_LIGHT = 'A29BFE';
const DARK_BG = '1A1D27';
const GRAY = '666666';
const LIGHT_BG = 'F5F5FA';
const WHITE = 'FFFFFF';
const BLACK = '000000';

const border = { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0 };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ heading: level, children: [new TextRun(text)] });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    ...opts,
    children: [new TextRun({ size: 22, font: 'Arial', ...opts.run, text })]
  });
}

function bulletItem(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    spacing: { after: 60 },
    children: [new TextRun({ size: 22, font: 'Arial', text })]
  });
}

function numberItem(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'numbers', level },
    spacing: { after: 60 },
    children: [new TextRun({ size: 22, font: 'Arial', text })]
  });
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, size: 28, bold: true, font: 'Arial', color: ACCENT })]
  });
}

function subTitle(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, size: 24, bold: true, font: 'Arial', color: '333333' })]
  });
}

function headerCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: ACCENT, type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: 'center',
    children: [new Paragraph({ children: [new TextRun({ text, size: 20, bold: true, font: 'Arial', color: WHITE })] })]
  });
}

function dataCell(text, width, opts = {}) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: opts.shade ? { fill: LIGHT_BG, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text: String(text), size: 20, font: 'Arial', ...opts.run })] })]
  });
}

function makeTable(headers, rows, colWidths) {
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({ children: headers.map((h, i) => headerCell(h, colWidths[i])) }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((cell, ci) => dataCell(cell, colWidths[ci], { shade: ri % 2 === 1 }))
      }))
    ]
  });
}

function spacer(h = 200) {
  return new Paragraph({ spacing: { after: h }, children: [] });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: 'Arial', color: DARK_BG },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 30, bold: true, font: 'Arial', color: '333333' },
        paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: 'Arial', color: ACCENT },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: 'bullets', levels: [
        { level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
      ]},
      { reference: 'numbers', levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      ]},
    ]
  },
  sections: [
    // ====== COVER PAGE ======
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      children: [
        spacer(2000),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
          children: [new TextRun({ text: 'iSens PC Bang Manager', size: 56, bold: true, font: 'Arial', color: ACCENT })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: 'Product Requirements Document (PRD)', size: 32, font: 'Arial', color: GRAY })] }),
        spacer(400),
        new Paragraph({ alignment: AlignmentType.CENTER, border: { top: { style: BorderStyle.SINGLE, size: 2, color: ACCENT } },
          spacing: { before: 300, after: 300 }, children: [] }),
        spacer(200),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: 'Version 1.0', size: 24, font: 'Arial', color: GRAY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: '2026.03.19', size: 24, font: 'Arial', color: GRAY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: '\u00A9 iSens League', size: 22, font: 'Arial', color: GRAY })] }),
      ]
    },

    // ====== MAIN CONTENT ======
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      headers: {
        default: new Header({ children: [
          new Paragraph({ alignment: AlignmentType.RIGHT, border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: ACCENT } },
            children: [new TextRun({ text: 'iSens PC Bang Manager - PRD v1.0', size: 16, font: 'Arial', color: GRAY, italics: true })] })
        ]})
      },
      footers: {
        default: new Footer({ children: [
          new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: 'Page ', size: 16, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GRAY })] })
        ]})
      },
      children: [
        // 1. 개요
        heading('1. 제품 개요', HeadingLevel.HEADING_1),

        subTitle('1.1 제품 비전'),
        para('iSens PC Bang Manager는 PC방(인터넷 카페) 운영자를 위한 통합 관리 솔루션으로, 실시간 PC 상태 모니터링, 좌석 배치 관리, 가동률 분석, 세리머니(이벤트 효과) 관리 등 PC방 운영에 필요한 모든 기능을 단일 웹 인터페이스에서 제공합니다.'),

        subTitle('1.2 제품 목표'),
        bulletItem('PC방 내 모든 PC의 실시간 상태(ON/OFF, 온도, 경고)를 시각적으로 모니터링'),
        bulletItem('도면 기반 직관적 좌석 배치 및 구역 관리'),
        bulletItem('시간대별 가동률 분석을 통한 운영 효율화'),
        bulletItem('세리머니(게임 이벤트 효과) 설정 및 통계 관리'),
        bulletItem('다중 매장 통합 관리 지원'),

        subTitle('1.3 기술 스택'),
        makeTable(
          ['항목', '상세'],
          [
            ['프론트엔드', 'HTML5, CSS3, Vanilla JavaScript'],
            ['디자인 시스템', '다크 테마, CSS Custom Properties, Pretendard 폰트'],
            ['아이콘', 'Remix Icon (CDN)'],
            ['차트', 'Canvas API 기반 자체 구현'],
            ['배포', '단일 HTML 파일 (Single File Application)'],
          ],
          [3000, 6360]
        ),

        spacer(),

        // 2. 사용자 관리
        heading('2. 사용자 인증 및 관리', HeadingLevel.HEADING_1),

        subTitle('2.1 로그인'),
        para('사용자는 아이디와 비밀번호로 시스템에 접근합니다. 로그인 화면에는 회사 로고가 표시되며, 비밀번호 표시/숨김 토글과 로그인 상태 유지 체크박스를 제공합니다.'),
        bulletItem('아이디/비밀번호 인증'),
        bulletItem('비밀번호 표시/숨김 토글'),
        bulletItem('로그인 상태 유지 (세션 스토리지)'),
        bulletItem('회원가입 페이지 링크'),

        subTitle('2.2 회원가입'),
        para('신규 사용자는 개인정보를 입력하여 가입을 신청합니다. 관리자 승인 후 로그인이 가능합니다.'),
        makeTable(
          ['필드', '필수 여부', '설명'],
          [
            ['회원 구분', '필수', '드롭다운 선택'],
            ['아이디', '필수', '고유 식별자'],
            ['비밀번호', '필수', '강도 표시 (약/보통/강)'],
            ['비밀번호 확인', '필수', '일치 검증'],
            ['이름', '필수', '실명'],
            ['휴대폰 번호', '필수', '연락처'],
            ['E-mail', '필수', '이메일 주소'],
            ['소속', '선택', '소속 부서/지점'],
            ['직급', '선택', '직급/직책'],
          ],
          [2400, 1500, 5460]
        ),

        spacer(),
        subTitle('2.3 개인정보관리'),
        para('사이드바 하단 프로필 영역에서 설정 메뉴를 통해 개인정보를 수정할 수 있습니다.'),
        bulletItem('계정 정보: 아이디(읽기전용), 비밀번호 변경'),
        bulletItem('개인 정보: 이름, 휴대폰 번호, E-mail 수정'),
        bulletItem('소속 정보: 소속, 직급, 회원 구분(읽기전용), 상태(읽기전용), 메모'),

        new PageBreak(),

        // 3. PC 관리 (메인 화면)
        heading('3. PC 관리 (메인 대시보드)', HeadingLevel.HEADING_1),

        subTitle('3.1 화면 구성'),
        para('PC 관리는 시스템의 메인 화면으로, 선택된 매장의 모든 PC 상태를 실시간으로 모니터링합니다.'),

        bulletItem('상단바: 매장 선택 드롭다운, PC 관리/가동률 탭, 상태 요약, 뷰 토글, 구역/편집 버튼'),
        bulletItem('좌측 사이드바: 메뉴 네비게이션, 프로필 영역'),
        bulletItem('메인 영역: 3가지 뷰(도면/구역/목록)로 PC 좌석 표시'),
        bulletItem('우측 패널: 선택된 PC의 상세 정보'),

        subTitle('3.2 상태 표시 체계'),
        makeTable(
          ['상태', '조건', '색상', '설명'],
          [
            ['가동(ON)', 'PC 전원 ON', '#00CEC9 (Cyan)', '정상 가동 중인 PC'],
            ['주의(Warning)', 'CPU 70-85°C 또는 GPU 65-80°C', '#FDCB6E (Yellow)', '온도 주의 필요'],
            ['경고(Error)', 'CPU 85°C+ 또는 GPU 80°C+', '#FF6B6B (Red)', '즉시 조치 필요'],
            ['미사용(OFF)', 'PC 전원 OFF', '#6B7185 (Gray)', '비가동 PC'],
          ],
          [1800, 2600, 2400, 2560]
        ),

        spacer(),
        para('상단 요약에서는 가동 PC 수에 주의/경고 PC가 포함됨을 괄호로 표시하며, 미사용 PC는 구분선으로 분리하여 직관적으로 표현합니다.', { run: { italics: true, color: GRAY } }),

        subTitle('3.3 도면 뷰 (Floor Plan View)'),
        para('매장 도면 위에 PC 좌석을 시각적으로 배치하여 물리적 위치를 확인할 수 있습니다.'),
        bulletItem('구역별 색상 코딩된 좌석 카드'),
        bulletItem('PC 번호, 상태 인디케이터, CPU 온도 표시'),
        bulletItem('줌 인/아웃 컨트롤 (50%~200%)'),
        bulletItem('좌석 호버 시 상세 툴팁 (IP, CPU, GPU, RAM)'),
        bulletItem('좌석 클릭 시 우측 상세 패널 오픈'),

        subTitle('3.4 구역 뷰 (Zone View)'),
        para('구역별로 그룹화된 카드 레이아웃으로 PC를 표시합니다.'),
        bulletItem('구역 헤더에 구역명과 좌석 수 표시'),
        bulletItem('각 구역 내 좌석을 가로 정렬로 배치'),
        bulletItem('편집 모드에서 구역 간 좌석 드래그&드롭 이동'),

        subTitle('3.5 목록 뷰 (List View)'),
        para('테이블 형태로 모든 PC의 상세 정보를 한눈에 파악할 수 있습니다.'),
        makeTable(
          ['컬럼', '고정 여부', '설명'],
          [
            ['PC번호', '고정 (left:0)', 'PC 식별자, 클릭 시 상세 패널'],
            ['구역', '고정 (left:80px)', '소속 구역명, 구역 색상 표시'],
            ['상태', '스크롤', 'ON/OFF 상태 뱃지'],
            ['IP', '스크롤', 'IP 주소'],
            ['CPU/GPU 모델', '스크롤', '하드웨어 사양'],
            ['CPU°/GPU°', '스크롤', '현재 온도 (임계값 초과 시 색상 경고)'],
            ['버전/업데이트', '스크롤', 'iSensManager 버전, 마지막 업데이트'],
          ],
          [2400, 2000, 4960]
        ),
        spacer(),
        para('페이지네이션: 우측 정렬, 페이지당 20~200개(20개 단위 선택), 상하단 동일하게 표시'),

        new PageBreak(),

        subTitle('3.6 PC 상세 정보 패널'),
        para('좌석 클릭 시 우측에 슬라이드되는 상세 정보 패널입니다.'),
        bulletItem('기본 정보: PC번호, 구역, IP, 버전, RAM, 장치 수'),
        bulletItem('CPU 섹션: 모델명, 온도 게이지, 현재/최고/경고 횟수'),
        bulletItem('GPU 섹션: 모델명, 온도 게이지, 현재/최고/경고 횟수'),
        bulletItem('연결 장치: Mouse, Keyboard, Headset 등 태그 표시'),
        bulletItem('24시간 가동률: 시간대별 바 차트, 우측에 평균 가동률(%) 표시'),
        bulletItem('하단 고정 툴바: 이전/수정/다음 버튼'),

        subTitle('3.7 편집 모드'),
        para('편집 버튼 클릭 시 활성화되며, 좌석 배치 및 구역 관리를 수정할 수 있습니다.'),
        bulletItem('구역 카드 드래그&드롭으로 순서 변경'),
        bulletItem('좌석 드래그로 구역 간 이동'),
        bulletItem('구역 크기 조절 (리사이즈 핸들)'),
        bulletItem('Shift+클릭으로 다중 선택'),
        bulletItem('우클릭 컨텍스트 메뉴 (수정/삭제)'),
        bulletItem('좌석 추가/삭제'),
        bulletItem('PC 관리, 가동률 탭 모두에서 편집 가능'),

        new PageBreak(),

        // 4. 가동률 탭
        heading('4. 가동률 분석', HeadingLevel.HEADING_1),

        subTitle('4.1 개요'),
        para('PC 관리 탭과 동일한 3가지 뷰(도면/구역/목록)에서 시간대별 PC 가동률을 시각적으로 분석합니다. 상단 탭 전환으로 PC 관리와 가동률 화면을 빠르게 전환할 수 있습니다.'),

        subTitle('4.2 시각화 방식'),
        para('단일 퍼플 컬러의 명도/채도 그라데이션으로 가동률을 6단계로 표현하여, 직관적이고 일관된 시각적 경험을 제공합니다.'),
        makeTable(
          ['가동률 범위', '배경 투명도', '시각적 특성'],
          [
            ['0% (미사용)', 'rgba(107,113,133,0.06)', '거의 투명한 회색'],
            ['1~20%', 'rgba(108,92,231,0.08)', '매우 연한 퍼플'],
            ['21~40%', 'rgba(108,92,231,0.14)', '연한 퍼플'],
            ['41~60%', 'rgba(108,92,231,0.22)', '중간 퍼플'],
            ['61~80%', 'rgba(108,92,231,0.32)', '진한 퍼플'],
            ['81~100%', 'rgba(108,92,231,0.45)', '가장 진한 퍼플'],
          ],
          [2400, 3460, 3500]
        ),

        spacer(),
        subTitle('4.3 상단 통계'),
        bulletItem('현재 시간대 표시 (예: 13~14시)'),
        bulletItem('평균 가동률 (%) 표시'),
        bulletItem('데이터 기준 시간 (예: 2026-03-19 13:00 기준)'),
        bulletItem('6단계 색상 범례'),

        subTitle('4.4 목록 뷰 특화'),
        para('가동률 목록 뷰에서는 24시간(00-01 ~ 23-24) 시간대별 컬럼이 추가되며, 각 셀이 가동률에 따라 색상 코딩됩니다. PC번호와 구역 컬럼은 좌측 고정됩니다.'),

        new PageBreak(),

        // 5. 매장 관리
        heading('5. 매장 관리', HeadingLevel.HEADING_1),

        subTitle('5.1 매장 목록'),
        para('등록된 모든 매장을 테이블 형태로 조회하고 관리합니다.'),
        makeTable(
          ['컬럼', '고정 여부', '설명'],
          [
            ['매장코드', '고정', '자동 생성 코드 (S1001~)'],
            ['매장명', '고정', '매장 이름'],
            ['대표자명', '스크롤', '매장 대표자'],
            ['가맹점주', '스크롤', '가맹점 점주'],
            ['본사 SV', '스크롤', '본사 담당 SV'],
            ['소속', '스크롤', '소속 지역/부서'],
            ['등록PC', '스크롤', '등록된 PC 수'],
            ['사업자등록번호', '스크롤', '사업자번호'],
            ['CPU/GPU 알람기준온도', '스크롤', '온도 경고 임계값'],
            ['추가 알람수신메일', '스크롤', '알림 수신 이메일'],
          ],
          [2400, 1600, 5360]
        ),

        spacer(),
        subTitle('5.2 매장 등록'),
        para('매장 등록 모달에서 새 매장을 추가합니다. 도면 이미지 업로드(PNG, JPG, SVG)도 함께 지원합니다.'),

        subTitle('5.3 매장 정보 수정'),
        para('매장 목록에서 행 클릭 시 수정 모달이 표시되며, 매장 정보를 수정하거나 삭제할 수 있습니다.'),

        new PageBreak(),

        // 6. PC 등록/수정
        heading('6. PC 등록/수정', HeadingLevel.HEADING_1),

        subTitle('6.1 PC 목록'),
        para('선택된 매장의 모든 PC를 관리합니다. 체크박스/PC상태/PC번호/IP 컬럼은 좌측 고정입니다.'),

        subTitle('6.2 주요 기능'),
        bulletItem('전체 PC / 가동 PC / 미사용 PC 필터'),
        bulletItem('전체 선택/해제 체크박스 (헤더)'),
        bulletItem('선택PC 체크: 선택된 PC의 상태 일괄 확인'),
        bulletItem('선택PC 삭제: 선택된 PC 일괄 삭제'),
        bulletItem('PC등록: 새 PC 등록 모달'),
        bulletItem('IP수정: IP 일괄 수정 모달'),

        subTitle('6.3 PC 세부내용 수정 모달'),
        para('편집 가능 필드와 읽기 전용 필드를 구분하여 표시합니다.'),
        makeTable(
          ['필드', '수정 가능', '설명'],
          [
            ['PC번호', 'O', 'PC 식별 번호'],
            ['구역', 'O', '드롭다운 선택'],
            ['IP', 'O', 'IP 주소'],
            ['iSensManager Ver.', 'O', '소프트웨어 버전'],
            ['RAM 용량', 'O', 'GB 단위'],
            ['연결 장치', 'O', '콤마 구분 입력'],
            ['CPU/GPU 사양', 'O', '하드웨어 모델명'],
            ['CPU/GPU 온도', 'X', '읽기 전용 (실시간 데이터)'],
            ['CPU/GPU 최고 온도', 'X', '읽기 전용'],
            ['CPU/GPU 알람 횟수', 'X', '읽기 전용'],
          ],
          [2800, 1200, 5360]
        ),

        new PageBreak(),

        // 7. 세리머니 관리
        heading('7. 세리머니 관리', HeadingLevel.HEADING_1),

        subTitle('7.1 개요'),
        para('세리머니는 게임 이벤트(승리, 킬, MVP 등) 발생 시 팀룸의 Shelly IoT 장치를 통해 조명/음향 효과를 실행하는 기능입니다. 통계 탭과 설정 탭으로 구분됩니다.'),

        subTitle('7.2 통계 탭'),
        bulletItem('기간 내 세리머니 실행수/중단수 카드 (대형 숫자 표시)'),
        bulletItem('시간/일/주/월 필터 토글'),
        bulletItem('조회 기간 날짜 범위 선택'),
        bulletItem('가장 많이 실행/중단한 세리머니 TOP 5 테이블'),
        bulletItem('가장 많이 실행/중단한 음원 TOP 5 테이블'),
        bulletItem('Canvas 기반 실행/중단 추이 라인 차트'),

        subTitle('7.3 설정 탭'),
        para('좌측 사이드바 + 우측 콘텐츠 레이아웃으로 팀룸별 세리머니를 설정합니다.'),

        para('좌측 사이드바:', { run: { bold: true } }),
        bulletItem('팀룸 목록 (추가/삭제 지원)'),
        bulletItem('음원 관리 (추가/수정/삭제 지원)'),

        para('우측 콘텐츠 (선택된 팀룸):', { run: { bold: true } }),
        bulletItem('기본 설정: 매장명, 팀룸명, ShellyIP, DeviceId, 연결 상태'),
        bulletItem('세레모니 설정: 6개 슬롯, 각각 세리머니명/음원 선택/물리 버튼 토글'),
        bulletItem('사운드 설정: 볼륨 강제 조정 토글, 볼륨 슬라이더 (0~100)'),

        new PageBreak(),

        // 8. 구역 관리
        heading('8. 구역 관리', HeadingLevel.HEADING_1),

        subTitle('8.1 기본 구역'),
        makeTable(
          ['구역명', '색상 코드', '용도'],
          [
            ['FPS존', '#6C5CE7 (Purple)', 'FPS 게임 전용 구역'],
            ['LOL존', '#E17055 (Orange)', 'LOL 게임 전용 구역'],
            ['VIP존', '#FDCB6E (Yellow)', 'VIP 고객 전용'],
            ['팀룸', '#FF6348 (Red)', '단체 이용 팀룸'],
            ['멀티존', '#636E72 (Gray)', '일반 멀티 게임'],
            ['FC ONLINE존', '#00B894 (Green)', 'FC ONLINE 전용'],
            ['커플존', '#FD79A8 (Pink)', '커플 전용 구역'],
            ['프렌즈존', '#00CEC9 (Cyan)', '친구 그룹 구역'],
            ['퍼스트클래스존', '#F9CA24 (Gold)', '프리미엄 구역'],
            ['덴탈존', '#74B9FF (Blue)', '치과/의료 제휴 구역'],
          ],
          [2400, 2800, 4160]
        ),

        spacer(),
        subTitle('8.2 구역 관리 기능'),
        bulletItem('구역 추가: 이름, 색상(프리셋 팔레트 또는 커스텀) 지정'),
        bulletItem('구역 수정: 이름 및 색상 변경'),
        bulletItem('구역 삭제: 확인 후 삭제, 소속 좌석 미배정 처리'),
        bulletItem('구역 패널: 좌측 슬라이드 패널에서 전체 구역 관리'),

        new PageBreak(),

        // 9. 실시간 데이터
        heading('9. 실시간 데이터 업데이트', HeadingLevel.HEADING_1),

        subTitle('9.1 온도 시뮬레이션'),
        para('8초 간격으로 가동 중인 PC의 온도 데이터를 갱신합니다.'),
        bulletItem('CPU 온도: 30~95°C 범위, ±2~5°C 랜덤 변동'),
        bulletItem('GPU 온도: 30~90°C 범위, ±2~5°C 랜덤 변동'),
        bulletItem('최고 온도 자동 갱신'),
        bulletItem('경고 횟수 자동 누적 (CPU 85°C 초과 시)'),
        bulletItem('상태 자동 변경 (온도 임계값 기반)'),

        subTitle('9.2 갱신 대상'),
        bulletItem('도면/구역/목록 뷰의 좌석 상태'),
        bulletItem('상단 통계바의 가동/주의/경고/미사용 수'),
        bulletItem('열린 상태의 상세 정보 패널'),

        new PageBreak(),

        // 10. UI/UX 디자인 시스템
        heading('10. 디자인 시스템', HeadingLevel.HEADING_1),

        subTitle('10.1 색상 팔레트'),
        makeTable(
          ['변수명', '색상 코드', '용도'],
          [
            ['--bg-primary', '#0F1117', '메인 배경'],
            ['--bg-secondary', '#1A1D27', '사이드바, 패널 배경'],
            ['--bg-tertiary', '#242736', '카드, 입력 필드 배경'],
            ['--accent', '#6C5CE7', '주요 강조색 (버튼, 링크)'],
            ['--accent-light', '#A29BFE', '보조 강조색'],
            ['--success', '#00CEC9', '성공, 가동 상태'],
            ['--warning', '#FDCB6E', '주의 상태'],
            ['--danger', '#FF6B6B', '경고, 에러, 삭제'],
            ['--text-primary', '#E8EAED', '주요 텍스트'],
            ['--text-secondary', '#9AA0B0', '보조 텍스트'],
            ['--text-muted', '#6B7185', '비활성 텍스트'],
          ],
          [2400, 2000, 4960]
        ),

        spacer(),
        subTitle('10.2 타이포그래피'),
        bulletItem('기본 폰트: Pretendard (한글 최적화)'),
        bulletItem('폴백: -apple-system, sans-serif'),
        bulletItem('기본 크기: 12px (본문), 10px (라벨), 17px (제목)'),

        subTitle('10.3 컴포넌트'),
        bulletItem('버튼: .btn (기본), .btn-primary (강조), .btn-danger (삭제), .btn-success (추가)'),
        bulletItem('입력: .form-input (다크 배경, 보더, 포커스 하이라이트)'),
        bulletItem('카드: .bg-card 배경, border-radius 10px, border 1px'),
        bulletItem('모달: 오버레이 + 중앙 패널, 헤더/바디/푸터 구분'),
        bulletItem('토스트: 하단 중앙, 2.5초 자동 사라짐'),
        bulletItem('토글 스위치: 44x24px, 슬라이딩 노브'),

        spacer(400),

        // 부록
        heading('부록: 데이터 모델', HeadingLevel.HEADING_1),

        subTitle('A. Seat (좌석) 객체'),
        makeTable(
          ['속성', '타입', '설명'],
          [
            ['id', 'number', '고유 식별자'],
            ['pcNumber', 'string', 'PC 표시 번호 (PC001~)'],
            ['zone', 'string', '소속 구역 ID'],
            ['status', 'string', 'on / off / warning / error'],
            ['isOn', 'boolean', '전원 상태'],
            ['ip', 'string', 'IP 주소'],
            ['cpuModel / gpuModel', 'string', 'CPU/GPU 모델명'],
            ['ram', 'number', 'RAM 용량 (GB)'],
            ['cpuTemp / gpuTemp', 'number', '현재 온도'],
            ['cpuMax / gpuMax', 'number', '최고 기록 온도'],
            ['cpuWarn / gpuWarn', 'number', '경고 발생 횟수'],
            ['devices', 'string[]', '연결 장치 목록'],
            ['utilization', 'number[24]', '24시간 가동률 배열 (0~100)'],
            ['lastUpdate', 'string', '마지막 업데이트 시각'],
          ],
          [2800, 1600, 4960]
        ),

        spacer(),
        subTitle('B. Zone (구역) 객체'),
        makeTable(
          ['속성', '타입', '설명'],
          [
            ['id', 'string', '구역 고유 ID'],
            ['name', 'string', '구역 이름'],
            ['color', 'string', '구역 색상 (hex)'],
          ],
          [2800, 1600, 4960]
        ),

        spacer(),
        subTitle('C. TeamRoom (팀룸/세리머니) 객체'),
        makeTable(
          ['속성', '타입', '설명'],
          [
            ['id', 'number', '고유 식별자'],
            ['name', 'string', '팀룸 이름'],
            ['shellyIp', 'string', 'Shelly 장치 IP'],
            ['deviceId', 'string', '장치 고유 ID'],
            ['connected', 'boolean', '연결 상태'],
            ['ceremonies', 'object[6]', '세리머니 슬롯 (이름, 음원, 활성화)'],
            ['volumeForce', 'boolean', '볼륨 강제 조정 여부'],
            ['volume', 'number', '볼륨 값 (0~100)'],
          ],
          [2800, 1600, 4960]
        ),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/Users/isens/Claude/pc-manager/iSens_PC_Bang_Manager_PRD_v1.0.docx', buffer);
  console.log('PRD document created successfully!');
});
