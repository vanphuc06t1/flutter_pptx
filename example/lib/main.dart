import 'dart:io';

import 'package:flutter/material.dart';
import 'package:flutter_pptx/flutter_pptx.dart';
import 'package:path_provider/path_provider.dart';

import 'download/download.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();

  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData.light(useMaterial3: true).copyWith(
        colorScheme: ColorScheme.fromSeed(
          seedColor: Colors.blue,
        ),
      ),
      home: const MyHomePage(title: 'Presentation Example'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  Future<FlutterPowerPoint> createPresentation() async {
    final pres = FlutterPowerPoint();

    pres.addTitleSlide(
      title: 'Slide one'.toTextValue(),
    );

    pres.addTitleAndPhotoSlide(
      title: 'Slide two'.toTextValue(),
      image: ImageReference(
        path: 'assets/images/sample_gif.gif',
        name: 'Sample Gif',
      ),
    );

    pres.addTitleAndPhotoAltSlide(
      title: 'Slide three'.toTextValue(),
      image: ImageReference(
        path: 'assets/images/sample_jpg.jpg',
        name: 'Sample Jpg',
      ),
    );

    pres
        .addTitleAndBulletsSlide(
          title: 'Slide three'.toTextValue(),
          bullets: [
            'Bullet 1',
            'Bullet 2',
            'Bullet 3',
            'Bullet 4',
          ].map((e) => e.toTextValue()).toList(),
        )
        .speakerNotes = TextValue.uniform('This is a note!');

    pres
        .addBulletsSlide(
          bullets: [
            'Bullet 1',
            'Bullet 2',
            'Bullet 3',
            'Bullet 4',
          ].map((e) => e.toTextValue()).toList(),
        )
        .speakerNotes = TextValue.singleLine([
      TextItem('This '),
      TextItem('is ', isBold: true),
      TextItem('a ', isUnderline: true),
      TextItem('note!'),
    ]);

    pres.addTitleBulletsAndPhotoSlide(
      title: 'Slide five'.toTextValue(),
      image: ImageReference(
        path: 'assets/images/sample_jpg.jpg',
        name: 'Sample Jpg',
      ),
      bullets: [
        'Bullet 1',
        'Bullet 2',
        'Bullet 3',
        'Bullet 4',
      ].map((e) => e.toTextValue()).toList(),
    );

    pres
        .addSectionSlide(
          section: 'Section 1'.toTextValue(),
        )
        .speakerNotes = TextValue.multiLine([
      TextValueLine(values: [
        TextItem('This '),
        TextItem('is ', isBold: true),
        TextItem('a ', isUnderline: true),
        TextItem('note 1!'),
      ]),
      TextValueLine(values: [
        TextItem('This '),
        TextItem('is ', isBold: true),
        TextItem('a ', isUnderline: true),
        TextItem('note 2!'),
      ]),
    ]);

    pres.addTitleOnlySlide(
      title: 'Title 1'.toTextValue(),
      subtitle: 'Subtitle 1'.toTextValue(),
    );

    pres.addAgendaSlide(
      title: 'Title 1'.toTextValue(),
      subtitle: 'Subtitle 1'.toTextValue(),
      topics: 'Topics 1'.toTextValue(),
    );

    pres.addStatementSlide(
      statement: 'Statement 1'.toTextValue(),
    );

    pres.addBigFactSlide(
      fact: 'Title 1'.toTextLine(),
      information: 'Fact 1'.toTextValue(),
    );

    pres.addQuoteSlide(
      quote: 'Quote 1'.toTextLine(),
      attribution: 'Attribution 1'.toTextValue(),
    );

    pres.addPhoto3UpSlide(
      image1: ImageReference(
        path: 'assets/images/sample_gif.gif',
        name: 'Sample Gif',
      ),
      image2: ImageReference(
        path: 'assets/images/sample_jpg.jpg',
        name: 'Sample Jpg',
      ),
      image3: ImageReference(
        path: 'assets/images/sample_png.png',
        name: 'Sample Png',
      ),
    );

    pres.addPhotoSlide(
      image: ImageReference(
        path: 'assets/images/sample_gif.gif',
        name: 'Sample Gif',
      ),
    );

    pres.addBlankSlide();

    pres.addBlankSlide().background.color = '000000';

    pres.addBlankSlide().background.image = ImageReference(
      path: 'assets/images/sample_gif.gif',
      name: 'Sample Gif',
    );

    await pres.addWidgetSlide(
      (size) => Center(
        child: Container(
          padding: const EdgeInsets.all(30.0),
          decoration: BoxDecoration(
            border: Border.all(color: Colors.blueAccent, width: 5.0),
            color: Colors.redAccent,
          ),
          child: const Text("This is an invisible widget"),
        ),
      ),
    );

    pres.showSlideNumbers = true;

    return pres;
  }

  Future<FlutterPowerPoint> createPresentationWithShapes() async {
    final pres = FlutterPowerPoint();

    ///XML example
    const String templateXML = r'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="object 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="604515"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="5B5863"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="object 3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="820801"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="object 4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="1037087"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="object 5"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="1253370"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="object 6"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="1469647"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="7" name="object 7"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="1685938"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="8" name="object 8"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="1902216"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="9" name="object 9"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="2118493"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="10" name="object 10"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="2334785"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="11" name="object 11"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="2551062"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="12" name="object 12"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="2767353"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="13" name="object 13"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="2983631"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="14" name="object 14"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="3199922"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="15" name="object 15"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="3416186"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="16" name="object 16"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="3632463"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="17" name="object 17"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="3848768"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="18" name="object 18"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="4065046"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="19" name="object 19"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="4281323"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="20" name="object 20"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="4497614"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="21" name="object 21"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="4713892"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="22" name="object 22"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="4930169"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="23" name="object 23"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="5146474"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="24" name="object 24"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="5362752"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="25" name="object 25"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="5579015"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="26" name="object 26"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="5795320"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="27" name="object 27"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="6011598"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="28" name="object 28"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="6227875"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="29" name="object 29"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="6444153"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="30" name="object 30"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="35343" y="6660444"/><a:ext cx="5441315" cy="0"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="5441315"><a:moveTo><a:pt x="5441239" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="0"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="3207"><a:solidFill><a:srgbClr val="2D2C31"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:grpSp><p:nvGrpSpPr><p:cNvPr id="31" name="object 31"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="599513" y="521903"/><a:ext cx="859790" cy="436880"/><a:chOff x="599513" y="521903"/><a:chExt cx="859790" cy="436880"/></a:xfrm></p:grpSpPr><p:pic><p:nvPicPr><p:cNvPr id="32" name="object 32"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="599513" y="607387"/><a:ext cx="106630" cy="229233"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:sp><p:nvSpPr><p:cNvPr id="33" name="object 33"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="685673" y="530001"/><a:ext cx="193040" cy="420370"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="193040" h="420369"><a:moveTo><a:pt x="193013" y="450"/></a:moveTo><a:lnTo><a:pt x="190763" y="1350"/></a:lnTo><a:lnTo><a:pt x="184239" y="0"/></a:lnTo><a:lnTo><a:pt x="179515" y="5848"/></a:lnTo><a:lnTo><a:pt x="159512" y="48007"/></a:lnTo><a:lnTo><a:pt x="147427" y="78756"/></a:lnTo><a:lnTo><a:pt x="140373" y="96281"/></a:lnTo><a:lnTo><a:pt x="123670" y="136606"/></a:lnTo><a:lnTo><a:pt x="105279" y="179967"/></a:lnTo><a:lnTo><a:pt x="86298" y="223328"/></a:lnTo><a:lnTo><a:pt x="67486" y="265000"/></a:lnTo><a:lnTo><a:pt x="49602" y="302203"/></a:lnTo><a:lnTo><a:pt x="41230" y="319317"/></a:lnTo><a:lnTo><a:pt x="33742" y="335187"/></a:lnTo><a:lnTo><a:pt x="16597" y="375332"/></a:lnTo><a:lnTo><a:pt x="1961" y="414570"/></a:lnTo><a:lnTo><a:pt x="0" y="420222"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="16197"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="34" name="object 34"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="806701" y="669475"/><a:ext cx="100556" cy="155671"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:pic><p:nvPicPr><p:cNvPr id="35" name="object 35"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId4" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="944825" y="693771"/><a:ext cx="105954" cy="100107"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:sp><p:nvSpPr><p:cNvPr id="36" name="object 36"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1075750" y="581742"/><a:ext cx="108585" cy="294640"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="108584" h="294640"><a:moveTo><a:pt x="107980" y="0"/></a:moveTo><a:lnTo><a:pt x="106856" y="2924"/></a:lnTo><a:lnTo><a:pt x="104381" y="6522"/></a:lnTo><a:lnTo><a:pt x="101232" y="17546"/></a:lnTo><a:lnTo><a:pt x="98859" y="26941"/></a:lnTo><a:lnTo><a:pt x="96255" y="38383"/></a:lnTo><a:lnTo><a:pt x="93101" y="51553"/></a:lnTo><a:lnTo><a:pt x="78201" y="100555"/></a:lnTo><a:lnTo><a:pt x="64788" y="139023"/></a:lnTo><a:lnTo><a:pt x="49097" y="179431"/></a:lnTo><a:lnTo><a:pt x="32394" y="220008"/></a:lnTo><a:lnTo><a:pt x="14173" y="262188"/></a:lnTo><a:lnTo><a:pt x="5821" y="281063"/></a:lnTo><a:lnTo><a:pt x="0" y="294245"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="16197"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="37" name="object 37"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId5" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1224223" y="674424"/><a:ext cx="234862" cy="127551"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:grpSp><p:pic><p:nvPicPr><p:cNvPr id="38" name="object 38"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId6" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1820807" y="341486"/><a:ext cx="1687190" cy="629068"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>''';

    pres.addXMLCustomSlide(
      templateXML: templateXML
    );

    pres.showSlideNumbers = true;

    return pres;
  }

  Future<void> downloadPresentation(FlutterPowerPoint pres) async {
    final bytes = await pres.save();
    if (bytes == null) return;
    downloadFile('presentation.pptx', bytes);
  }

  @override
  void initState() {
    super.initState();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text(widget.title),
      ),
      body: Center(
          child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          ElevatedButton(
            onPressed: () async {
              final pres = await createPresentation();
              await downloadPresentation(pres);
            },
            child: const Text('Download Presentation'),
          ),
          ElevatedButton(
            onPressed: () async {
              final pres = await createPresentationWithShapes();
              Directory appDocDir = await getApplicationDocumentsDirectory();
              String appDocumentDir = appDocDir.path;
              print('ENV: appDocumentDir: $appDocumentDir');

              var exportFile = File("${appDocumentDir}/presentation.pptx");
              final bytes = await pres.save();
              if (bytes != null) {
                exportFile.writeAsBytes(bytes.toList(),
                    mode: FileMode.write, flush: true);
              }
            },
            child: const Text('Export Presentation'),
          ),
        ],
      )),
    );
  }
}
