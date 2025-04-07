// docx/image/constants.dart

class JPEG_MARKER_CODE {
 static const List<int> TEM = [0x01];
 static const List<int> DHT = [0xC4];
 // ... etc ...
 static const List<int> SOI = [0xD8];
 static const List<int> EOI = [0xD9];
 static const List<int> SOS = [0xDA];
 static const List<int> APP0 = [0xE0];
 static const List<int> APP1 = [0xE1];
 // ... etc ...

 static final List<List<int>> STANDALONE_MARKERS = [TEM, SOI, EOI, /* ...RST0-7...*/];
 static final List<List<int>> SOF_MARKER_CODES = [/*...SOF0-SOFF...*/];

 static const Map<String, String> markerNames = {
    // Usando representação de string hexadecimal para as chaves
    '00': 'UNKNOWN',
    'C0': 'SOF0',
    'C2': 'SOF2',
    // ... etc ...
 };

 static bool isStandalone(List<int> markerCode) {
   // Implementação para verificar se markerCode está em STANDALONE_MARKERS
   return false; // Placeholder
 }
}

class MIME_TYPE {
 static const String BMP = "image/bmp";
 static const String GIF = "image/gif";
 static const String JPEG = "image/jpeg";
 static const String PNG = "image/png";
 static const String TIFF = "image/tiff";
}

class PNG_CHUNK_TYPE {
 static const String IHDR = "IHDR";
 static const String pHYs = "pHYs";
 static const String IEND = "IEND";
}

class TIFF_FLD_TYPE {
 static const int BYTE = 1;
 static const int ASCII = 2;
 // ... etc ...
 static const int RATIONAL = 5;

 static const Map<int, String> fieldTypeNames = {
    1: "BYTE",
    2: "ASCII char",
    // ... etc ...
 };
}
typedef TIFF_FLD = TIFF_FLD_TYPE;

class TIFF_TAG {
 static const int IMAGE_WIDTH = 0x0100;
 static const int IMAGE_LENGTH = 0x0101;
 static const int X_RESOLUTION = 0x011A;
 static const int Y_RESOLUTION = 0x011B;
 static const int RESOLUTION_UNIT = 0x0128;

 static const Map<int, String> tagNames = {
    0x00FE: "NewSubfileType",
    0x0100: "ImageWidth",
    // ... etc ...
 };
}