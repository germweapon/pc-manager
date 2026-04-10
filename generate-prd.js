const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, TableOfContents, TabStopType, TabStopPosition } = require('docx');

// ── Color palette ──
const C = {
  primary: '1B2A4A',
  accent: '2E75B6',
  accentLight: 'D5E8F0',
  headerBg: '1B2A4A',
  headerText: 'FFFFFF',
  altRow: 'F2F7FB',
  border: 'B4C6D9',
  textDark: '1A1A1A',
  textMid: '444444',
  textLight: '666666',
  success: '27AE60',
  warning: 'F39C12',
  danger: 'E74C3C',
};

const border = { style: BorderStyle.SINGLE, size: 1, color: C.border };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0 };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

const PAGE_W = 12240;
const MARGIN = 1440;
const CONTENT_W = PAGE_W - 2 * MARGIN; // 9360

// ── Helpers ──
function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ heading: level, children: [new TextRun(text)] });
}

function para(text, opts = {}) {
  const runs = [];
  if (typeof text === 'string') {
    runs.push(new TextRun({ text, size: opts.size || 22, color: opts.color || C.textDark, font: 'Arial', ...opts.run }));
  } else {
    text.forEach(t => runs.push(new TextRun({ size: 22, color: C.textDark, font: 'Arial', ...t })));
  }
  return new Paragraph({
    spacing: { after: opts.after !== undefined ? opts.after : 160, before: opts.before || 0, line: opts.line || 300 },
    alignment: opts.align || AlignmentType.LEFT,
    children: runs,
    ...(opts.para || {}),
  });
}

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: C.headerBg, type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: 'center',
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, size: 20, color: C.headerText, font: 'Arial' })]
    })],
  });
}

function cell(text, width, opts = {}) {
  const children = typeof text === 'string'
    ? [new Paragraph({
        alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({ text, size: 20, color: opts.color || C.textDark, font: 'Arial', bold: opts.bold })]
      })]
    : (Array.isArray(text) ? text : [text]);
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    verticalAlign: opts.verticalAlign || 'center',
    children,
  });
}

function simpleTable(headers, rows, colWidths) {
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({ children: headers.map((h, i) => headerCell(h, colWidths[i])) }),
      ...rows.map((row, ri) =>
        new TableRow({
          children: row.map((c, ci) => {
            if (typeof c === 'object' && c._cell) return c._cell;
            return cell(String(c), colWidths[ci], { shading: ri % 2 === 1 ? C.altRow : undefined });
          })
        })
      ),
    ],
  });
}

function spacer(h = 200) {
  return new Paragraph({ spacing: { after: h } });
}

// ── Build Document ──
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: 'Arial', color: C.primary },
        paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.accent, space: 8 } }
        },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: C.accent },
        paragraph: { spacing: { before: 280, after: 180 }, outlineLevel: 1 },
      },
      {
        id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: C.primary },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: 'bullets2',
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: '\u2013', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
        ],
      },
      {
        reference: 'numbers',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: 'numbers2',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: 'numbers3',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: 'numbers4',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [
    // ═══════════════════════════════════════
    // COVER PAGE
    // ═══════════════════════════════════════
    {
      properties: {
        page: {
          size: { width: PAGE_W, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: [
        spacer(2400),
        // Title block
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: 'iSens Manager', size: 56, bold: true, color: C.accent, font: 'Arial' })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: 'PC\uBC29 \uC88C\uC11D \uB9E4\uB2C8\uC800', size: 40, bold: true, color: C.primary, font: 'Arial' })],
        }),
        spacer(200),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 2, color: C.accent, space: 12 }, bottom: { style: BorderStyle.SINGLE, size: 2, color: C.accent, space: 12 } },
          spacing: { after: 40 },
          children: [new TextRun({ text: 'Product Requirements Document (PRD)', size: 28, color: C.textMid, font: 'Arial' })],
        }),
        spacer(600),
        // Meta info table
        new Table({
          width: { size: 5000, type: WidthType.DXA },
          columnWidths: [2000, 3000],
          rows: [
            ['문서 버전', 'v1.0'],
            ['작성일', '2026-03-16'],
            ['작성자', 'iSens 개발팀'],
            ['보안 등급', '사내 기밀'],
            ['상태', '초안 (Draft)'],
          ].map(([k, v]) =>
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: k, size: 20, bold: true, color: C.textLight, font: 'Arial' })] })],
                }),
                new TableCell({
                  borders: noBorders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 80, right: 80 },
                  children: [new Paragraph({ children: [new TextRun({ text: v, size: 20, color: C.textDark, font: 'Arial' })] })],
                }),
              ],
            })
          ),
        }),
      ],
    },

    // ═══════════════════════════════════════
    // TOC + MAIN CONTENT
    // ═══════════════════════════════════════
    {
      properties: {
        page: {
          size: { width: PAGE_W, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.accent, space: 6 } },
            children: [
              new TextRun({ text: 'iSens Manager PRD', size: 16, color: C.textLight, font: 'Arial', italics: true }),
              new TextRun({ text: '\tv1.0', size: 16, color: C.textLight, font: 'Arial', italics: true }),
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 1, color: C.border, space: 6 } },
            children: [
              new TextRun({ text: '\u00A9 iSens  |  Confidential  |  Page ', size: 16, color: C.textLight, font: 'Arial' }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.textLight, font: 'Arial' }),
            ],
          })],
        }),
      },
      children: [
        // TOC
        heading('목차', HeadingLevel.HEADING_1),
        new TableOfContents('Table of Contents', { hyperlink: true, headingStyleRange: '1-3' }),
        new Paragraph({ children: [new PageBreak()] }),

        // ── 1. 개요 ──
        heading('1. 제품 개요'),

        heading('1.1 프로젝트 배경', HeadingLevel.HEADING_2),
        para('아이센스 블랙라벨 PC방은 전국 다수 매장을 운영하는 프리미엄 PC방 프랜차이즈로, 매장별로 100대 이상의 고성능 PC를 관리해야 합니다. 기존 텍스트 기반의 PC 관리 시스템은 매장 내 물리적 좌석 배치와 PC 상태를 직관적으로 파악하기 어려우며, 온도 이상이나 장비 고장 등의 문제에 대한 즉각적 대응이 지연되는 한계가 있었습니다.'),
        para('이러한 문제를 해결하기 위해, 매장 좌석 도면 기반의 시각적 PC 모니터링 및 관리 시스템인 "iSens Manager"를 개발합니다.'),

        heading('1.2 제품 비전', HeadingLevel.HEADING_2),
        para([
          { text: '"매장 도면 위에서 모든 PC를 한눈에."', bold: true, italics: true, size: 24, color: C.accent },
        ]),
        para('iSens Manager는 실제 매장의 좌석 배치도를 웹 상에서 인터랙티브하게 재구성하여, PC별 실시간 상태(온도, 가동 여부, 경고)를 시각적으로 모니터링하고, 좌석의 추가/수정/삭제/자동 정렬까지 직관적으로 관리할 수 있는 올인원 PC 매니저 웹 애플리케이션입니다.'),

        heading('1.3 대상 사용자', HeadingLevel.HEADING_2),
        simpleTable(
          ['사용자 유형', '역할', '주요 니즈'],
          [
            ['매장 관리자', 'PC 상태 실시간 모니터링', '온도 이상 PC 즉시 파악, 빠른 대응'],
            ['본사 운영팀', '전체 매장 현황 총괄 관리', '매장별 가동률, 장비 상태 통계'],
            ['기술 지원팀', 'PC 하드웨어 원격 점검', 'CPU/GPU 온도, 장치 연결 상태 확인'],
            ['매장 점주', '영업 현황 파악', '좌석 가동률, 인기 구역 분석'],
          ],
          [2200, 3000, 4160],
        ),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 2. 핵심 기능 요구사항 ──
        heading('2. 핵심 기능 요구사항'),

        heading('2.1 좌석 도면 뷰 (Floor Plan View)', HeadingLevel.HEADING_2),
        para('매장의 실제 좌석 배치를 웹 상에서 시각적으로 재현하는 핵심 화면입니다.', { after: 80 }),

        heading('2.1.1 도면 이미지 등록', HeadingLevel.HEADING_3),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '매장별 좌석 안내도 이미지(PNG, JPG, SVG) 업로드 지원', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '업로드된 이미지를 배경으로 표시 (투명도 조절 가능)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '이미지 기반으로 좌석 블록을 HTML 직사각형으로 재구성', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '최대 10MB, 지원 형식: PNG, JPG, SVG', size: 22, font: 'Arial' })] }),

        heading('2.1.2 좌석 블록 표시', HeadingLevel.HEADING_3),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '각 좌석을 직사각형 블록으로 표시', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '좌석 번호 표시 (예: 1, 2, 3...)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: 'PC 상태에 따른 색상 코딩 (사용중/미사용/주의/경고)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '실시간 CPU 온도 표시 (가동 중인 PC)', size: 22, font: 'Arial' })] }),

        // Status color table
        simpleTable(
          ['상태', '색상', '조건', '설명'],
          [
            ['사용중 (ON)', '청록색 (#00cec9)', 'PC 전원 ON + 정상 범위', '정상 가동 상태'],
            ['주의 (Warning)', '노란색 (#fdcb6e)', 'CPU > 70°C 또는 GPU > 65°C', '온도 주의 필요'],
            ['경고 (Error)', '빨간색 (#ff6b6b)', 'CPU > 85°C 또는 GPU > 80°C', '즉시 조치 필요'],
            ['미사용 (OFF)', '회색', 'PC 전원 OFF', '비가동 상태'],
          ],
          [1600, 2400, 2800, 2560],
        ),
        spacer(),

        heading('2.1.3 좌석 호버 툴팁', HeadingLevel.HEADING_3),
        para('마우스를 좌석 블록 위에 올리면 즉시 핵심 정보를 보여주는 플로팅 툴팁을 표시합니다.'),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: 'PC번호, 구역명, 상태(ON/OFF)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: 'IP 주소, CPU 모델 요약', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: 'CPU/GPU 온도 (경고 시 색상 강조)', size: 22, font: 'Arial' })] }),

        heading('2.1.4 좌석 상세 패널', HeadingLevel.HEADING_3),
        para('좌석 클릭 시 우측에 슬라이드 인(slide-in) 상세 패널이 열리며, PC의 전체 정보를 표시합니다.'),
        simpleTable(
          ['정보 섹션', '항목', '비고'],
          [
            ['기본 정보', 'PC번호, 구역, IP, 버전, RAM, 장치 수', '6개 항목 그리드 배치'],
            ['CPU 정보', '모델명, 현재 온도, 최고 온도, 경고 횟수', '온도 게이지 바 시각화'],
            ['GPU 정보', '모델명, 현재 온도, 최고 온도, 경고 횟수', '온도 게이지 바 시각화'],
            ['연결 장치', 'Mouse, Keyboard, Headset, Speaker 등', '태그(칩) 형태 표시'],
            ['가동률', '24시간 시간대별 가동률 차트', '막대 차트 시각화'],
            ['업데이트', '마지막 데이터 업데이트 시각', '타임스탬프 표시'],
          ],
          [2000, 4000, 3360],
        ),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 2.2 좌석 편집 ──
        heading('2.2 좌석 편집 기능', HeadingLevel.HEADING_2),
        para('관리자가 좌석 레이아웃을 자유롭게 구성하고 수정할 수 있는 편집 모드입니다.'),

        heading('2.2.1 편집 모드 진입/종료', HeadingLevel.HEADING_3),
        new Paragraph({ numbering: { reference: 'numbers', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '상단 "편집" 버튼 또는 사이드바 "좌석 편집" 메뉴로 편집 모드 진입', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'numbers', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '편집 모드 진입 시 상단에 편집 툴바 표시 (안내 메시지 + 기능 버튼)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'numbers', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '좌석 블록에 편집 인디케이터(보라색 점) 표시, 드래그 가능 상태', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'numbers', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '"편집 종료" 버튼으로 모드 해제 및 변경사항 자동 저장', size: 22, font: 'Arial' })] }),

        heading('2.2.2 좌석 추가', HeadingLevel.HEADING_3),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '"좌석 추가" 버튼 클릭 후 도면의 빈 공간 클릭으로 새 좌석 배치', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '자동 PC 번호 할당 (기존 최대 번호 + 1)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '좌석 배치 시 그리드 스냅 (10px 단위)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '기본 구역: 멀티존, 이후 우클릭 메뉴로 변경 가능', size: 22, font: 'Arial' })] }),

        heading('2.2.3 좌석 수정', HeadingLevel.HEADING_3),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '우클릭 컨텍스트 메뉴 > "좌석 번호 수정"으로 번호 변경', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '우클릭 > "구역 변경"으로 소속 구역 변경 (모달에서 구역 선택)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '편집 모드에서 좌석 드래그로 위치 이동 (5px 단위 스냅)', size: 22, font: 'Arial' })] }),

        heading('2.2.4 좌석 삭제', HeadingLevel.HEADING_3),
        para([
          { text: '단일 삭제: ', bold: true },
          { text: '우클릭 컨텍스트 메뉴 > "좌석 삭제" 선택 시 삭제 확인 모달 표시' },
        ], { after: 80 }),
        para([
          { text: '다중 삭제: ', bold: true },
          { text: 'Shift+클릭 또는 드래그 범위 선택으로 여러 좌석을 다중 선택한 후 "선택 삭제" 버튼으로 일괄 삭제' },
        ], { after: 80 }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '삭제 확인 모달에서 삭제 대상 PC번호를 칩(chip) 형태로 미리보기 표시', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '"이 작업은 되돌릴 수 없습니다" 경고 메시지 포함', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '다중 선택 시 선택된 좌석 수를 편집 툴바에 실시간 표시', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '선택 영역 드래그: 빈 공간에서 드래그하면 빨간 점선 박스로 범위 표시', size: 22, font: 'Arial' })] }),

        heading('2.2.5 좌석 위치 자동 정렬', HeadingLevel.HEADING_3),
        para('편집 툴바의 "자동 정렬" 버튼 클릭 시 5가지 정렬 옵션을 제공하는 모달이 표시됩니다.'),
        simpleTable(
          ['정렬 유형', '동작 설명', '적용 대상'],
          [
            ['그리드 정렬', '모든 좌석을 균일한 그리드 포인트(56px 단위)에 스냅, 겹침 자동 해소', '전체 또는 선택된 좌석'],
            ['구역별 정렬', '각 구역의 기준점(좌상단)을 유지하며 구역 내 좌석을 행/열로 균등 재배치', '전체 또는 선택된 좌석'],
            ['빈틈 없이 정렬', '좌석 간 간격을 최소화하여 밀집 배치, 구역별로 그룹핑 유지', '전체 또는 선택된 좌석'],
            ['수평 정렬', '선택된 좌석의 평균 Y좌표로 모두 정렬, X축 방향 균등 배치', '선택된 좌석 (2개 이상)'],
            ['수직 정렬', '선택된 좌석의 평균 X좌표로 모두 정렬, Y축 방향 균등 배치', '선택된 좌석 (2개 이상)'],
          ],
          [2000, 4800, 2560],
        ),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 2.3 구역 관리 ──
        heading('2.3 구역(Zone) 관리', HeadingLevel.HEADING_2),
        para('PC방 내 좌석을 게임 장르, 서비스 등급 등 목적에 따라 구역으로 분류합니다.'),

        heading('2.3.1 기본 제공 구역', HeadingLevel.HEADING_3),
        simpleTable(
          ['구역명', '색상 코드', '용도'],
          [
            ['FPS존', '#6c5ce7 (보라)', 'FPS 게임 전용 고성능 구역'],
            ['LOL존', '#e17055 (오렌지)', 'LOL/AOS 게임 구역'],
            ['VIP존', '#fdcb6e (골드)', '프리미엄 좌석 구역'],
            ['팀룸', '#e17055 (레드)', '단체 이용 전용 구역'],
            ['커플존', '#fd79a8 (핑크)', '2인 전용 좌석 구역'],
            ['프렌즈존', '#00cec9 (청록)', '그룹 이용 구역'],
            ['퍼스트클래스존', '#f9ca24 (옐로우)', '최상위 프리미엄 구역'],
            ['멀티존', '#636e72 (그레이)', '일반 다목적 구역'],
            ['FC ONLINE존', '#00b894 (그린)', 'FC Online 전용 구역'],
            ['덴탈존', '#74b9ff (스카이블루)', '특수 구역'],
          ],
          [2600, 3200, 3560],
        ),
        spacer(),

        heading('2.3.2 구역 관리 기능', HeadingLevel.HEADING_3),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '좌측 구역 패널에서 전체 구역 목록 및 좌석 수 확인', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '신규 구역 추가: 구역명 + 10가지 프리셋 색상 중 선택', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '도면에서 구역별 영역을 반투명 배경색으로 시각적 구분', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '구역 영역 좌상단에 구역명 라벨 표시', size: 22, font: 'Arial' })] }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 2.4 PC 목록 뷰 ──
        heading('2.4 PC 목록 뷰 (Table View)', HeadingLevel.HEADING_2),
        para('전체 PC를 테이블 형태로 조회하는 데이터 중심 뷰입니다.'),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '상단 뷰 토글(도면/목록)로 전환', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '표시 컬럼: PC상태, PC번호, IP, 구역, CPU 모델, RAM, CPU 온도, CPU Max, GPU 모델, GPU 온도, GPU Max, 버전, 업데이트 시각', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '온도 값은 위험 수준에 따라 색상 강조 (빨강/노랑)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '행 클릭 시 도면 뷰와 동일한 상세 패널 열기', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '헤더 고정(sticky header)으로 스크롤 시에도 컬럼명 유지', size: 22, font: 'Arial' })] }),

        // ── 2.5 가동률 조회 ──
        heading('2.5 가동률 조회', HeadingLevel.HEADING_2),
        para('PC별 시간대별 가동률을 조회하는 기능입니다.'),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '날짜 범위 선택 (시작일 ~ 종료일)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '매장 선택 필터', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '시간대별 가동률 표시 (00~01, 01~02, ... 23~24)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '전체 가동률 요약 (컬럼 헤더에 시간대별 평균 가동률 표시)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '상세 패널 내 24시간 미니 차트로 개별 PC 가동 패턴 시각화', size: 22, font: 'Arial' })] }),

        // ── 2.6 실시간 모니터링 ──
        heading('2.6 실시간 모니터링', HeadingLevel.HEADING_2),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '5초 주기 온도 데이터 자동 갱신 (CPU/GPU)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '상태 변경 시 좌석 블록 색상 자동 업데이트', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '상단 바에 전체 현황 요약 (가동 수, 미사용 수, 경고 수)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '최고 온도(Max) 자동 기록 및 경고 횟수 누적', size: 22, font: 'Arial' })] }),

        // ── 2.7 다중 매장 ──
        heading('2.7 다중 매장 지원', HeadingLevel.HEADING_2),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '사이드바 하단 매장 선택 드롭다운으로 매장 전환', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '매장별 독립적 도면 레이아웃 및 구역 설정', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '지원 매장: 수원경희대점, 강남역점, 홍대입구점, 부산서면점, 대전둔산점 등', size: 22, font: 'Arial' })] }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 3. UI/UX 설계 ──
        heading('3. UI/UX 설계 방향'),

        heading('3.1 디자인 시스템', HeadingLevel.HEADING_2),

        heading('3.1.1 색상 체계 (Dark Theme)', HeadingLevel.HEADING_3),
        simpleTable(
          ['용도', '색상', 'Hex Code'],
          [
            ['배경 (Primary)', '다크 네이비', '#0f1117'],
            ['배경 (Secondary)', '진한 회색', '#1a1d27'],
            ['카드 배경', '블루 그레이', '#1e2130'],
            ['테두리', '어두운 보라 회색', '#2e3144'],
            ['텍스트 (Primary)', '밝은 회색', '#e8eaed'],
            ['텍스트 (Secondary)', '중간 회색', '#9aa0b0'],
            ['Accent', '보라', '#6c5ce7'],
            ['Success', '청록', '#00cec9'],
            ['Warning', '노란색', '#fdcb6e'],
            ['Danger', '빨간색', '#ff6b6b'],
          ],
          [3000, 3000, 3360],
        ),
        spacer(),

        heading('3.1.2 타이포그래피', HeadingLevel.HEADING_3),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '기본 폰트: Pretendard (가변 폰트, CDN 로딩)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '폴백: -apple-system, sans-serif', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '아이콘: Remix Icon (CDN)', size: 22, font: 'Arial' })] }),

        heading('3.1.3 컴포넌트', HeadingLevel.HEADING_3),
        simpleTable(
          ['컴포넌트', '용도', '특징'],
          [
            ['좌석 블록 (Seat)', '도면 위 PC 표시', '52x44px, 상태별 색상, 호버 효과, 온도 표시'],
            ['상세 패널 (Detail)', 'PC 상세 정보', '우측 360px 슬라이드-인, 섹션별 정보 구성'],
            ['툴팁 (Tooltip)', '빠른 정보 확인', '호버 시 PC 핵심 정보 6개 항목 표시'],
            ['컨텍스트 메뉴', '좌석 관리 액션', '우클릭 메뉴: 번호 수정, 구역 변경, 삭제'],
            ['모달 (Modal)', '확인/입력 UI', '블러 배경 오버레이, 슬라이드-인 애니메이션'],
            ['토스트 (Toast)', '피드백 알림', '하단 중앙, 2.5초 자동 닫힘'],
            ['편집 툴바', '편집 모드 안내', '상단 고정, 안내 텍스트 + 기능 버튼'],
          ],
          [2400, 2400, 4560],
        ),

        new Paragraph({ children: [new PageBreak()] }),

        heading('3.2 레이아웃 구조', HeadingLevel.HEADING_2),
        para('전체 레이아웃은 3-컬럼 구조로 설계됩니다.'),
        simpleTable(
          ['영역', '너비', '역할'],
          [
            ['사이드바 (Sidebar)', '260px (고정)', '내비게이션, 매장 선택, 메뉴'],
            ['메인 (Main)', '가변 (flex: 1)', '도면 뷰 / 테이블 뷰 + 상단바'],
            ['구역 패널 (Zone)', '240px (토글)', '구역 목록 및 관리 (접고 펼치기)'],
            ['상세 패널 (Detail)', '360px (토글)', 'PC 상세 정보 (클릭 시 표시)'],
          ],
          [2600, 2400, 4360],
        ),

        heading('3.3 인터랙션 패턴', HeadingLevel.HEADING_2),
        simpleTable(
          ['인터랙션', '트리거', '결과'],
          [
            ['좌석 클릭', '좌석 블록 좌클릭', '상세 패널 오픈 + 좌석 하이라이트'],
            ['좌석 호버', '마우스 오버', '툴팁 표시 (커서 추적)'],
            ['우클릭 메뉴', '좌석 블록 우클릭', '컨텍스트 메뉴 (수정/구역/삭제)'],
            ['좌석 드래그', '편집 모드 + 드래그', '좌석 위치 이동 (5px 스냅)'],
            ['다중 선택', 'Shift+클릭 / 드래그', '여러 좌석 선택 (빨간 테두리)'],
            ['줌', '줌 버튼 (+/-)', '도면 확대/축소 (50%~200%)'],
            ['뷰 전환', '토글 버튼', '도면 뷰 <-> 테이블 뷰'],
          ],
          [2200, 2800, 4360],
        ),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 4. 데이터 모델 ──
        heading('4. 데이터 모델'),

        heading('4.1 PC 좌석 데이터', HeadingLevel.HEADING_2),
        simpleTable(
          ['필드명', '타입', '설명', '예시'],
          [
            ['id', 'number', '좌석 고유 ID', '1, 2, 3...'],
            ['pcNumber', 'string', 'PC 식별 번호', 'PC001, PC002'],
            ['zone', 'string', '소속 구역 ID', 'fps, lol, vip'],
            ['x, y', 'number', '도면 위 좌표 (px)', 'x: 30, y: 30'],
            ['status', 'string', '현재 상태', 'on / off / warning / error'],
            ['isOn', 'boolean', '전원 상태', 'true / false'],
            ['ip', 'string', 'IP 주소', '125.130.6.141'],
            ['cpuModel', 'string', 'CPU 모델명', 'Intel Core i9-9900K'],
            ['gpuModel', 'string', 'GPU 모델명', 'NVIDIA GeForce RTX 3060'],
            ['ram', 'number', 'RAM 용량 (GB)', '16, 32'],
            ['cpuTemp', 'number', '현재 CPU 온도', '31 ~ 95'],
            ['gpuTemp', 'number', '현재 GPU 온도', '30 ~ 90'],
            ['cpuMax / gpuMax', 'number', '최고 기록 온도', '60 ~ 95'],
            ['cpuWarn / gpuWarn', 'number', '경고 발생 횟수', '0, 1, 2...'],
            ['devices', 'string[]', '연결된 주변기기 목록', '["Mouse", "Keyboard"]'],
            ['version', 'string', 'iSensManager 버전', '1.2.0.2'],
            ['utilization', 'number[]', '24시간 가동률 배열', '[0, 4.58, 33.64, ...]'],
            ['lastUpdate', 'string', '마지막 업데이트 시각', '2026-03-16 15:18:40'],
          ],
          [2200, 1400, 3000, 2760],
        ),

        new Paragraph({ children: [new PageBreak()] }),

        heading('4.2 구역 데이터', HeadingLevel.HEADING_2),
        simpleTable(
          ['필드명', '타입', '설명', '예시'],
          [
            ['id', 'string', '구역 고유 ID', 'fps, lol, vip'],
            ['name', 'string', '구역 표시명', 'FPS존, VIP존'],
            ['color', 'string', '구역 대표 색상 (hex)', '#6c5ce7'],
          ],
          [2200, 1400, 3000, 2760],
        ),

        spacer(),

        // ── 5. 기술 스택 ──
        heading('5. 기술 스택 및 아키텍처'),

        heading('5.1 현재 프로토타입 기술 스택', HeadingLevel.HEADING_2),
        simpleTable(
          ['구분', '기술', '비고'],
          [
            ['프론트엔드', 'HTML5 + CSS3 + Vanilla JS', '단일 파일 프로토타입'],
            ['스타일링', 'CSS Custom Properties (변수)', 'Dark Theme 기반'],
            ['폰트', 'Pretendard (CDN)', '가변 웹폰트'],
            ['아이콘', 'Remix Icon (CDN)', '벡터 아이콘'],
            ['백엔드 연동', '미구현 (Mock Data)', '향후 REST API 연동 예정'],
          ],
          [2200, 3400, 3760],
        ),
        spacer(),

        heading('5.2 향후 프로덕션 아키텍처 (권장)', HeadingLevel.HEADING_2),
        simpleTable(
          ['계층', '기술 스택', '역할'],
          [
            ['프론트엔드', 'React + TypeScript', 'SPA, 컴포넌트 기반 UI'],
            ['상태 관리', 'Zustand / Redux Toolkit', 'PC 상태, 좌석 레이아웃 관리'],
            ['실시간 통신', 'WebSocket / SSE', 'PC 온도 실시간 스트리밍'],
            ['백엔드 API', 'Node.js (Express) 또는 Spring Boot', 'REST API + WebSocket'],
            ['데이터베이스', 'PostgreSQL + Redis', '영구 저장 + 실시간 캐시'],
            ['모니터링 에이전트', 'iSensManager Agent (C++/C#)', 'PC별 설치, 5초 주기 데이터 전송'],
            ['인프라', 'Docker + AWS/NCP', '컨테이너 배포, 오토 스케일링'],
          ],
          [2200, 3800, 3360],
        ),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 6. 비기능 요구사항 ──
        heading('6. 비기능 요구사항'),

        heading('6.1 성능', HeadingLevel.HEADING_2),
        simpleTable(
          ['항목', '목표', '측정 기준'],
          [
            ['초기 로딩', '2초 이내', 'FCP (First Contentful Paint)'],
            ['실시간 갱신', '5초 이내', '온도 데이터 업데이트 주기'],
            ['도면 렌더링', '200대 이상 PC 동시 표시', '프레임 드롭 없이 60fps 유지'],
            ['줌/패닝', '지연 없음', '사용자 입력 후 16ms 이내 반영'],
            ['API 응답', '500ms 이내', 'PC 목록 조회, 좌석 정보 CRUD'],
          ],
          [2400, 3200, 3760],
        ),
        spacer(),

        heading('6.2 호환성', HeadingLevel.HEADING_2),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '브라우저: Chrome 90+, Edge 90+, Firefox 90+ (IE 미지원)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '해상도: 최소 1280x720, 권장 1920x1080', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: '디바이스: 데스크탑 전용 (모바일 지원은 Phase 2에서 검토)', size: 22, font: 'Arial' })] }),

        heading('6.3 보안', HeadingLevel.HEADING_2),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '관리자 계정 인증 (로그인 필수)', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 300 }, children: [new TextRun({ text: '역할 기반 접근 제어 (RBAC): 매장 관리자 / 본사 운영팀', size: 22, font: 'Arial' })] }),
        new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 160, line: 300 }, children: [new TextRun({ text: 'HTTPS 통신 필수, API 토큰 기반 인증', size: 22, font: 'Arial' })] }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── 7. 개발 로드맵 ──
        heading('7. 개발 로드맵'),

        simpleTable(
          ['Phase', '기간', '주요 목표', '산출물'],
          [
            ['Phase 1', '2026 Q2 (4주)', '프로토타입 완성 및 사용성 검증', 'HTML 프로토타입, 사용성 테스트 결과'],
            ['Phase 2', '2026 Q2~Q3 (8주)', 'MVP 개발 (React 전환 + 백엔드 구축)', 'React SPA + REST API + DB'],
            ['Phase 3', '2026 Q3 (4주)', '실시간 모니터링 연동', 'WebSocket 통신, Agent 연동'],
            ['Phase 4', '2026 Q4 (4주)', '다중 매장 지원 + 통계 기능', '매장 관리, 대시보드, 리포트'],
            ['Phase 5', '2027 Q1 (4주)', '고급 기능 + 안정화', '알림 시스템, 원격 제어, 모바일 지원'],
          ],
          [1200, 1800, 3200, 3160],
        ),

        spacer(300),

        // ── 8. 성공 지표 ──
        heading('8. 성공 지표 (KPI)'),

        simpleTable(
          ['지표', '목표', '측정 방법'],
          [
            ['온도 이상 대응 시간', '기존 대비 50% 단축', '경고 발생 ~ 조치 완료까지 시간'],
            ['매장 관리자 만족도', '4.5/5.0 이상', '분기별 설문 조사'],
            ['도면 뷰 사용률', '전체 접속의 70% 이상', '뷰 전환 로그 분석'],
            ['시스템 가용성', '99.5% 이상', '월간 다운타임 모니터링'],
            ['PC 장애 감소율', '전년 대비 30% 감소', '장애 접수 건수 비교'],
          ],
          [2400, 2800, 4160],
        ),

        spacer(300),

        // ── 9. 용어 정의 ──
        heading('9. 용어 정의'),

        simpleTable(
          ['용어', '정의'],
          [
            ['iSens Manager', 'PC방 좌석 도면 기반 관리 웹 애플리케이션의 제품명'],
            ['좌석 도면 (Floor Plan)', '매장 내 좌석 배치를 시각적으로 표현한 인터랙티브 뷰'],
            ['좌석 블록 (Seat Block)', '도면 위에 배치되는 개별 PC를 나타내는 직사각형 UI 요소'],
            ['구역 (Zone)', '게임 장르, 서비스 등급 등으로 분류된 좌석 그룹'],
            ['iSensManager Agent', '각 PC에 설치되어 하드웨어 정보를 수집/전송하는 모니터링 에이전트'],
            ['가동률', '특정 시간대에 PC가 사용된 비율 (%)'],
          ],
          [2800, 6560],
        ),

        spacer(400),

        // ── Footer ──
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 2, color: C.accent, space: 12 } },
          spacing: { before: 200 },
          children: [new TextRun({ text: '-- End of Document --', size: 20, color: C.textLight, font: 'Arial', italics: true })],
        }),
      ],
    },
  ],
});

// Generate
const outputPath = '/Users/isens/Claude/pc-manager/iSens_Manager_PRD_v1.0.docx';
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log('PRD generated:', outputPath);
  console.log('Size:', (buffer.length / 1024).toFixed(1), 'KB');
});
