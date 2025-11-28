import 'package:docx_dart/src/oxml/text/hyperlink.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/text/run.dart';
import 'package:docx_dart/src/types.dart';

/// Proxy object wrapping a `<w:hyperlink>` element.
class Hyperlink extends Parented {
  Hyperlink(this._hyperlink, ProvidesStoryPart parent)
      : _storyParent = parent,
        super(parent);

  final CT_Hyperlink _hyperlink;
  final ProvidesStoryPart _storyParent;

  /// The hyperlink URI (address portion). Empty string for internal jumps.
  String get address {
    final rId = _hyperlink.rId;
    if (rId == null) {
      return '';
    }
    final rel = part.rels[rId];
    return rel?.targetRef ?? '';
  }

    /// True when the hyperlink text is broken across page boundaries.
  bool get containsPageBreak =>
      _hyperlink.lastRenderedPageBreaks.isNotEmpty;

  /// Bookmark or fragment reference without the leading '#'. Empty when absent.
  String get fragment => _hyperlink.anchor ?? '';

  /// Sequence of runs contained in this hyperlink.
  List<Run> get runs =>
      _hyperlink.r_lst.map((r) => Run(r, _storyParent)).toList(growable: false);

  /// Textual content produced by concatenating the runs in this hyperlink.
  String get text => _hyperlink.text;

  /// Convenience accessor returning `address#fragment` when both portions exist.
  String get url {
    final base = address;
    final frag = fragment;
    if (base.isEmpty) {
      return '';
    }
    return frag.isEmpty ? base : '$base#$frag';
  }
}
