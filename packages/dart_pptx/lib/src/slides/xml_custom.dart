import 'package:json_annotation/json_annotation.dart';

import '../classes/slide.dart';

@JsonSerializable(createFactory: false)
class SlideXMLCustom extends Slide {
  SlideXMLCustom({
    super.name = 'XML Custom',
    required this.templateXML,
    super.slideNumber,
  });

  String templateXML;

  @override
  int get layoutId => 1;

  @override
  String get source => templateXML;

}