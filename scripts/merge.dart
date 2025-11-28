//merge.dart
import 'dart:io';

void main() {
  // Lista de diretórios a serem processados.
  final List<String> directoryPaths = [
    r'C:\MyDartProjects\docx_dart\python-docx\src\docx'
  ];

  // Lista de extensões de arquivo que serão incluídas.
  final List<String> extensions = ['.py'];

  // Define o diretório de saída.
  final String outputDirectoryPath = r'C:\MyDartProjects\docx_dart\output';

  // Cria o diretório de saída caso ele não exista.
  final outputDir = Directory(outputDirectoryPath);
  if (!outputDir.existsSync()) {
    outputDir.createSync(recursive: true);
  }

  // Define o arquivo de saída que irá conter o merge.
  final outputFile = File('$outputDirectoryPath/merged_files.py');

  // Limpa o arquivo de saída caso ele já exista.
  if (outputFile.existsSync()) {
    outputFile.writeAsStringSync('');
  }

  // Processa cada diretório da lista.
  for (var dirPath in directoryPaths) {
    final directory = Directory(dirPath);
    // Verifica se o diretório existe.
    if (!directory.existsSync()) {
      print('Diretório não encontrado: $dirPath');
      continue;
    }

    // Lista os arquivos do diretório.
    final List<FileSystemEntity> entities = directory.listSync(recursive: true);

    // Processa cada arquivo encontrado.
    for (var entity in entities) {
      if (entity is File) {
        // Verifica se o arquivo possui uma das extensões desejadas.
        if (extensions.any((ext) => entity.path.endsWith(ext))) {
          // Extrai o nome do arquivo.
          final fileName = entity.uri.pathSegments.last;
          // Lê o conteúdo do arquivo.
          final content = entity.readAsStringSync();

          // Escreve um comentário com o nome do arquivo seguido do seu conteúdo.
          outputFile.writeAsStringSync('# $fileName\n', mode: FileMode.append);
          outputFile.writeAsStringSync(content, mode: FileMode.append);
          outputFile.writeAsStringSync('\n\n', mode: FileMode.append);
        }
      }
    }
  }

  print('Merge completo! Arquivo "merged_files.py" criado com sucesso no diretório "$outputDirectoryPath".');
}
