part of excel_community;

/// Receives the bytes of a zip as [ZipStreamWriter] produces them.
abstract interface class ByteSink {
  void add(Uint8List bytes);
}

/// How a zip entry's content is stored.
enum ZipEntryMethod { stored, deflate }

/// Writes a zip archive to a [ByteSink] entry by entry while the content is
/// produced, never holding an entry, let alone the archive, in memory.
///
/// Every entry opened with [beginEntry] is written with a data descriptor
/// (its CRC and sizes follow the content), so the writer needs neither size
/// nor checksum up front; the central directory written by [finish] carries
/// the real values. An entry that may grow past 4 GiB is opened with
/// `large: true`, which reserves zip64 sizes in its local header; the central
/// directory and the end record switch to zip64 on their own when a size, an
/// offset or the entry count overflows.
///
/// ```dart
/// final zip = ZipStreamWriter(sink);
/// zip.beginEntry('xl/worksheets/sheet1.xml', large: true);
/// zip.write(chunk); // as often as needed
/// zip.endEntry();
/// zip.addEntry('[Content_Types].xml', utf8.encode(types));
/// zip.finish();
/// ```
class ZipStreamWriter {
  ZipStreamWriter(this._sink, {this.deflateLevel = DeflateLevel.defaultCompression});

  static const int _localHeaderSignature = 0x04034b50;
  static const int _dataDescriptorSignature = 0x08074b50;
  static const int _centralDirectorySignature = 0x02014b50;
  static const int _endOfCentralDirectorySignature = 0x06054b50;
  static const int _zip64EndOfCentralDirectorySignature = 0x06064b50;
  static const int _zip64LocatorSignature = 0x07064b50;
  static const int _zip64ExtraId = 0x0001;
  static const int _flagDataDescriptor = 0x0008;
  static const int _flagUtf8 = 0x0800;
  static const int _versionZip64 = 45;
  static const int _versionDeflate = 20;
  static const int _max32 = 0xFFFFFFFF;
  static const int _max16 = 0xFFFF;

  final ByteSink _sink;
  final int deflateLevel;
  final List<_ZipStreamEntry> _entries = [];

  int _position = 0;
  _ZipStreamEntry? _open;
  StreamDeflate? _deflate;
  int _dataStart = 0;
  bool _finished = false;

  /// Bytes handed to the sink so far.
  int get bytesWritten => _position;

  /// Opens [name]; the content follows through [write] until [endEntry].
  /// [large] reserves zip64 sizes for an entry that may pass 4 GiB.
  void beginEntry(
    String name, {
    ZipEntryMethod method = ZipEntryMethod.deflate,
    DateTime? modified,
    bool large = false,
  }) {
    _checkOpen(false);
    final entry = _ZipStreamEntry(
      name: name.replaceAll('\\', '/'),
      method: method,
      offset: _position,
      modified: modified ?? DateTime.now(),
      large: large,
      flags: _flagDataDescriptor | _flagUtf8,
    );
    _open = entry;
    _writeLocalHeader(entry);
    _dataStart = _position;
    _deflate = method == ZipEntryMethod.deflate ? StreamDeflate(level: deflateLevel) : null;
  }

  /// Appends [bytes] to the open entry.
  void write(Uint8List bytes) {
    final entry = _open;
    if (entry == null) throw StateError('No entry is open');
    if (bytes.isEmpty) return;
    entry.crc32 = getCrc32(bytes, entry.crc32);
    entry.uncompressedSize += bytes.length;
    final deflate = _deflate;
    if (deflate == null) {
      _emit(bytes);
    } else {
      deflate.addBytes(bytes);
      _emit(deflate.takeBytes());
    }
  }

  /// Closes the open entry: flushes the deflate stream and writes the data
  /// descriptor with the CRC and the sizes.
  void endEntry() {
    final entry = _open;
    if (entry == null) throw StateError('No entry is open');
    final deflate = _deflate;
    if (deflate != null) {
      deflate.finish();
      _emit(deflate.takeBytes());
    }
    entry.compressedSize = _position - _dataStart;
    if (!entry.large && (entry.compressedSize > _max32 || entry.uncompressedSize > _max32)) {
      throw StateError('Entry "${entry.name}" passed 4 GiB; open it with large: true');
    }
    final descriptor = _ByteWriter()
      ..u32(_dataDescriptorSignature)
      ..u32(entry.crc32);
    if (entry.large) {
      descriptor
        ..u64(entry.compressedSize)
        ..u64(entry.uncompressedSize);
    } else {
      descriptor
        ..u32(entry.compressedSize)
        ..u32(entry.uncompressedSize);
    }
    _emit(descriptor.bytes);
    _entries.add(entry);
    _open = null;
    _deflate = null;
  }

  /// Drops the open entry: its bytes stay in the stream (readers go by the
  /// central directory, which will not list it), so a failed content stream
  /// leaves the archive readable with the entry missing.
  void abandonEntry() {
    final entry = _open;
    if (entry == null) throw StateError('No entry is open');
    final deflate = _deflate;
    if (deflate != null) {
      deflate.finish();
      _emit(deflate.takeBytes());
    }
    entry.compressedSize = _position - _dataStart;
    final descriptor = _ByteWriter()
      ..u32(_dataDescriptorSignature)
      ..u32(entry.crc32);
    if (entry.large) {
      descriptor
        ..u64(entry.compressedSize)
        ..u64(entry.uncompressedSize);
    } else {
      descriptor
        ..u32(entry.compressedSize & _max32)
        ..u32(entry.uncompressedSize & _max32);
    }
    _emit(descriptor.bytes);
    _open = null;
    _deflate = null;
  }

  /// Writes a whole entry at once.
  void addEntry(
    String name,
    List<int> bytes, {
    ZipEntryMethod method = ZipEntryMethod.deflate,
    DateTime? modified,
  }) {
    beginEntry(name, method: method, modified: modified);
    write(bytes is Uint8List ? bytes : Uint8List.fromList(bytes));
    endEntry();
  }

  /// Writes an empty directory entry.
  void addDirectory(String name, {DateTime? modified}) {
    _checkOpen(false);
    final path = name.replaceAll('\\', '/');
    final entry = _ZipStreamEntry(
      name: path.endsWith('/') ? path : '$path/',
      method: ZipEntryMethod.stored,
      offset: _position,
      modified: modified ?? DateTime.now(),
      large: false,
      flags: _flagUtf8,
      isDirectory: true,
    );
    _writeLocalHeader(entry);
    _entries.add(entry);
  }

  /// Writes the central directory and the end record. Nothing can be added
  /// afterwards.
  void finish() {
    _checkOpen(false);
    if (_finished) throw StateError('finish() was already called');
    _finished = true;
    final directoryOffset = _position;
    for (final entry in _entries) {
      _writeCentralEntry(entry);
    }
    final directorySize = _position - directoryOffset;
    final needsZip64 = _entries.length > _max16 || directorySize > _max32 || directoryOffset > _max32;
    if (needsZip64) {
      final zip64Offset = _position;
      _emit(
        (_ByteWriter()
              ..u32(_zip64EndOfCentralDirectorySignature)
              ..u64(44)
              ..u16(_versionZip64)
              ..u16(_versionZip64)
              ..u32(0)
              ..u32(0)
              ..u64(_entries.length)
              ..u64(_entries.length)
              ..u64(directorySize)
              ..u64(directoryOffset))
            .bytes,
      );
      _emit(
        (_ByteWriter()
              ..u32(_zip64LocatorSignature)
              ..u32(0)
              ..u64(zip64Offset)
              ..u32(1))
            .bytes,
      );
    }
    _emit(
      (_ByteWriter()
            ..u32(_endOfCentralDirectorySignature)
            ..u16(0)
            ..u16(0)
            ..u16(_entries.length > _max16 ? _max16 : _entries.length)
            ..u16(_entries.length > _max16 ? _max16 : _entries.length)
            ..u32(directorySize > _max32 ? _max32 : directorySize)
            ..u32(directoryOffset > _max32 ? _max32 : directoryOffset)
            ..u16(0))
          .bytes,
    );
  }

  void _checkOpen(bool expected) {
    if (_finished) throw StateError('The archive is finished');
    if ((_open != null) != expected) {
      throw StateError(expected ? 'No entry is open' : 'Entry "${_open!.name}" is still open');
    }
  }

  void _writeLocalHeader(_ZipStreamEntry entry) {
    final name = utf8.encode(entry.name);
    final header = _ByteWriter()
      ..u32(_localHeaderSignature)
      ..u16(entry.large ? _versionZip64 : _versionDeflate)
      ..u16(entry.flags)
      ..u16(entry.methodCode)
      ..u16(entry.dosTime)
      ..u16(entry.dosDate)
      ..u32(0)
      ..u32(entry.large ? _max32 : 0)
      ..u32(entry.large ? _max32 : 0)
      ..u16(name.length)
      ..u16(entry.large ? 20 : 0)
      ..add(name);
    if (entry.large) {
      header
        ..u16(_zip64ExtraId)
        ..u16(16)
        ..u64(0)
        ..u64(0);
    }
    _emit(header.bytes);
  }

  void _writeCentralEntry(_ZipStreamEntry entry) {
    final name = utf8.encode(entry.name);
    final sizeOverflow = entry.uncompressedSize > _max32 || entry.compressedSize > _max32;
    final offsetOverflow = entry.offset > _max32;
    final extra = _ByteWriter();
    if (sizeOverflow || offsetOverflow) {
      final fields = _ByteWriter();
      if (sizeOverflow) {
        fields
          ..u64(entry.uncompressedSize)
          ..u64(entry.compressedSize);
      }
      if (offsetOverflow) fields.u64(entry.offset);
      extra
        ..u16(_zip64ExtraId)
        ..u16(fields.length)
        ..add(fields.bytes);
    }
    final record = _ByteWriter()
      ..u32(_centralDirectorySignature)
      ..u16(_versionZip64)
      ..u16(sizeOverflow || offsetOverflow || entry.large ? _versionZip64 : _versionDeflate)
      ..u16(entry.flags)
      ..u16(entry.methodCode)
      ..u16(entry.dosTime)
      ..u16(entry.dosDate)
      ..u32(entry.crc32)
      ..u32(sizeOverflow ? _max32 : entry.compressedSize)
      ..u32(sizeOverflow ? _max32 : entry.uncompressedSize)
      ..u16(name.length)
      ..u16(extra.length)
      ..u16(0)
      ..u16(0)
      ..u16(0)
      ..u32(entry.isDirectory ? 0x10 : 0)
      ..u32(offsetOverflow ? _max32 : entry.offset)
      ..add(name)
      ..add(extra.bytes);
    _emit(record.bytes);
  }

  void _emit(Uint8List bytes) {
    if (bytes.isEmpty) return;
    _sink.add(bytes);
    _position += bytes.length;
  }
}

class _ZipStreamEntry {
  final String name;
  final ZipEntryMethod method;
  final int offset;
  final DateTime modified;
  final bool large;
  final int flags;
  final bool isDirectory;
  int crc32 = 0;
  int compressedSize = 0;
  int uncompressedSize = 0;

  _ZipStreamEntry({
    required this.name,
    required this.method,
    required this.offset,
    required this.modified,
    required this.large,
    required this.flags,
    this.isDirectory = false,
  });

  int get methodCode => method == ZipEntryMethod.deflate && !isDirectory ? 8 : 0;

  int get dosTime => (modified.hour << 11) | (modified.minute << 5) | (modified.second >> 1);

  int get dosDate {
    final year = modified.year < 1980 ? 1980 : modified.year;
    return ((year - 1980) << 9) | (modified.month << 5) | modified.day;
  }
}

/// Little-endian record builder for the zip headers.
class _ByteWriter {
  final BytesBuilder _builder = BytesBuilder(copy: false);

  int get length => _builder.length;

  Uint8List get bytes => _builder.toBytes();

  void u16(int value) => _builder.add(Uint8List(2)..buffer.asByteData().setUint16(0, value, Endian.little));

  void u32(int value) => _builder.add(Uint8List(4)..buffer.asByteData().setUint32(0, value, Endian.little));

  /// Two 32-bit halves: dart2js has no 64-bit accessors, and the arithmetic
  /// split is exact for every size a zip can hold (below 2^53).
  void u64(int value) {
    final high = value ~/ 0x100000000;
    u32(value - high * 0x100000000);
    u32(high);
  }

  void add(List<int> value) => _builder.add(value);
}
