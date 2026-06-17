library docx_dart;

// --- Core API ---
export 'src/api.dart';
export 'src/document.dart';

// --- Units & Shared Types ---
export 'src/shared.dart'
    show Length, Inches, Cm, Mm, Pt, Emu, Twips, RGBColor;

// --- User Friendly Formatting API ---
export 'src/formatting.dart';

// --- Enumerations ---
export 'src/enum/text.dart'
    show
        WD_PARAGRAPH_ALIGNMENT,
        WD_ALIGN_PARAGRAPH,
        WD_UNDERLINE,
        WD_BREAK,
        WD_LINE_SPACING,
        WD_COLOR_INDEX;
export 'src/enum/section.dart'
    show WD_ORIENTATION, WD_SECTION, WD_SECTION_START, WD_HEADER_FOOTER;

// --- Text (Paragraph, Run, Font, ParagraphFormat) ---
export 'src/text/paragraph.dart' show Paragraph;
export 'src/text/run.dart' show Run;
export 'src/text/font_proxy.dart' show Font;
export 'src/text/parfmt.dart' show ParagraphFormat;

// --- Section ---
export 'src/section.dart' show Section, Sections;

// --- Table ---
export 'src/table.dart' show Table;

// --- Shapes ---
export 'src/shape.dart' show InlineShape, InlineShapes;

// --- Styles ---
export 'src/styles/style.dart' show CharacterStyle, ParagraphStyle;

// --- Settings ---
export 'src/settings.dart' show Settings;

// --- Core Properties ---
export 'src/opc/coreprops.dart' show CoreProperties;
