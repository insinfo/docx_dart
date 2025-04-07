

```
lib/
└── src/
    ├── exceptions.dart                       # exceptions.py (Core Exceptions)
    │
    ├── opc/
    │   ├── constants.dart                    # opc/constants.py
    │   ├── exceptions.dart                   # opc/exceptions.py
    │   ├── shared.dart                       # opc/shared.py
    │   ├── spec.dart                         # opc/spec.py
    │   └── packuri.dart                      # opc/packuri.py
    │
    ├── oxml/
    │   ├── exceptions.dart                   # oxml/exceptions.py
    │   ├── ns.dart                           # oxml/ns.py
    │   ├── xmlchemy.dart                     # oxml/xmlchemy.py (Base para mapeamento OXML)
    │   ├── simpletypes.dart                  # oxml/simpletypes.py (Tipos base XML)
    │   ├── shared.dart                       # oxml/shared.py (Tipos OXML compartilhados como CT_OnOff)
    │   ├── coreprops.dart                    # oxml/coreprops.py (CT_CoreProperties)
    │   ├── drawing.dart                      # oxml/drawing.py (CT_Drawing)
    │   ├── text/
    │   │   ├── font.dart       parei aqui    # oxml/text/font.py (CT_RPr, etc.)
    │   │   ├── run.dart                      # oxml/text/run.py (CT_R, CT_Text, etc.)
    │   │   ├── hyperlink.dart                # oxml/text/hyperlink.py (CT_Hyperlink)
    │   │   ├── pagebreak.dart                # oxml/text/pagebreak.py (CT_LastRenderedPageBreak)
    │   │   ├── parfmt.dart                   # oxml/text/parfmt.py (CT_PPr, etc.)
    │   │   └── paragraph.dart                # oxml/text/paragraph.py (CT_P)
    │   ├── shape.dart                        # oxml/shape.py (Shape CT_* types)
    │   ├── table.dart                        # oxml/table.py (Table CT_* types)
    │   ├── section.dart                      # oxml/section.py (Section CT_* types)
    │   ├── document.dart                     # oxml/document.py (CT_Document, CT_Body)
    │   ├── numbering.dart                    # oxml/numbering.py (Numbering CT_* types)
    │   ├── settings.dart                     # oxml/settings.py (CT_Settings)
    │   ├── styles.dart                       # oxml/styles.py (Styles CT_* types)
    │   ├── parser.dart                       # oxml/parser.py (Funções de análise/registro - adaptadas para package:xml)
    │   └── __init__.dart                     # oxml/__init__.py (Registro de elementos - adaptado)
    │
    ├── opc/
    │   ├── oxml.dart                         # opc/oxml.py (OPC CT_* types)
    │   ├── rel.dart                          # opc/rel.py (Relationships)
    │   ├── part.dart                         # opc/part.py (Part, XmlPart, PartFactory)
    │   ├── parts/
    │   │   └── coreprops.dart                # opc/parts/coreprops.py (CorePropertiesPart)
    │   ├── phys_pkg.dart                     # opc/phys_pkg.py (Leitura/Escrita Física - usando package:archive)
    │   ├── pkgreader.dart                    # opc/pkgreader.py (PackageReader)
    │   ├── pkgwriter.dart                    # opc/pkgwriter.py (PackageWriter)
    │   └── package.dart                      # opc/package.py (OpcPackage base)
    │
    ├── enum/
    │   ├── base.dart                         # enum/base.py
    │   ├── dml.dart                          # enum/dml.py
    │   ├── section.dart                      # enum/section.py
    │   ├── shape.dart                        # enum/shape.py
    │   ├── style.dart                        # enum/style.py
    │   ├── table.dart                        # enum/table.py
    │   └── text.dart                         # enum/text.py
    │
    ├── shared.dart                           # shared.py (Length, RGBColor, ElementProxy, Parented, etc.)
    │
    ├── image/
    │   ├── constants.dart                    # image/constants.py
    │   ├── exceptions.dart                   # image/exceptions.py
    │   ├── helpers.dart                      # image/helpers.py (StreamReader adaptado)
    │   ├── image.dart                        # image/image.py (BaseImageHeader, Image)
    │   ├── tiff.dart                         # image/tiff.py
    │   ├── bmp.dart                          # image/bmp.py
    │   ├── gif.dart                          # image/gif.py
    │   ├── jpeg.dart                         # image/jpeg.py
    │   ├── png.dart                          # image/png.py
    │   └── __init__.dart                     # image/__init__.py (SIGNATURES - adaptado)
    │
    ├── parts/
    │   ├── image.dart                        # parts/image.py (ImagePart)
    │   ├── story.dart                        # parts/story.py (StoryPart)
    │   ├── hdrftr.dart                       # parts/hdrftr.py (HeaderPart, FooterPart)
    │   ├── numbering.dart                    # parts/numbering.py (NumberingPart)
    │   ├── settings.dart                     # parts/settings.py (SettingsPart)
    │   ├── styles.dart                       # parts/styles.py (StylesPart)
    │   └── document.dart                     # parts/document.py (DocumentPart)
    │
    ├── package.dart                          # package.py (WordprocessingML Package)
    │
    ├── styles/
    │   ├── __init__.dart                     # styles/__init__.py (BabelFish)
    │   ├── latent.dart                       # styles/latent.py
    │   ├── style.dart                        # styles/style.py (StyleFactory, BaseStyle, etc.)
    │   └── styles.dart                       # styles/styles.py (Styles collection)
    │
    ├── dml/
    │   └── color.dart                        # color.py (Renomeado para dml/color.dart)
    │
    ├── drawing/
    │   └── __init__.dart                     # drawing/__init__.py
    │
    ├── text/
    │   ├── font.dart                         # text/font.py
    │   ├── tabstops.dart                     # text/tabstops.py
    │   ├── parfmt.dart                       # text/parfmt.py
    │   ├── run.dart                          # text/run.py
    │   ├── hyperlink.dart                    # text/hyperlink.py
    │   ├── pagebreak.dart                    # text/pagebreak.py
    │   └── paragraph.dart                    # text/paragraph.py
    │
    ├── blkcntnr.dart                         # blkcntnr.py (BlockItemContainer)
    │
    ├── table.dart                            # table.py
    │
    ├── section.dart                          # section.py
    │
    ├── settings.dart                         # settings.py
    │
    ├── shape.dart                            # shape.py
    │
    ├── document.dart                         # document.py (Document API class)
    │
    ├── api.dart                              # api.py (Função Document() de alto nível)
    │
    └── __init__.dart                         # __init__.py (Exportação principal e registro de PartFactory)
```