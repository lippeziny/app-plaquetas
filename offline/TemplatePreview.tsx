import React from 'react';
import { Image, Pressable, Text, View, useWindowDimensions } from 'react-native';
import type { SlidePositionSettings, TemplateSlide, TemplateVisual } from './pptx';

type PreviewItem = { code: string; name: string };

type Props = {
  slide: TemplateSlide;
  labels: PreviewItem[];
  adjustments?: SlidePositionSettings;
  activeSlot: number;
  activeField: 'code' | 'name';
  accent: string;
  onSelectField: (slot: number, field: 'code' | 'name') => void;
};

const EMU_PER_MM = 36000;

function addDelta(globalDelta?: { xMm: number; yMm: number }, slotDelta?: { xMm: number; yMm: number }) {
  return {
    x: ((globalDelta?.xMm || 0) + (slotDelta?.xMm || 0)) * EMU_PER_MM,
    y: ((globalDelta?.yMm || 0) + (slotDelta?.yMm || 0)) * EMU_PER_MM,
  };
}

function PreviewShape({
  visual,
  slide,
  labels,
  adjustments,
  activeSlot,
  activeField,
  accent,
  onSelectField,
  scale,
}: Props & { visual: TemplateVisual; scale: number }) {
  const codeIndex = visual.field === 'code' ? slide.codeFields.findIndex(field => field.id === visual.id) : -1;
  const nameIndex = visual.field === 'name' ? slide.nameFields.findIndex(field => field.id === visual.id) : -1;
  const slot = codeIndex >= 0 ? codeIndex : nameIndex;
  const field = codeIndex >= 0 ? 'code' : nameIndex >= 0 ? 'name' : null;
  const item = slot >= 0 ? labels[slot] : undefined;
  const globalDelta = field ? adjustments?.global?.[field] : undefined;
  const slotDelta = field && slot >= 0 ? adjustments?.slots?.[String(slot)]?.[field] : undefined;
  const delta = addDelta(globalDelta, slotDelta);
  const left = (visual.x + delta.x) * scale;
  const top = (visual.y + delta.y) * scale;
  const width = Math.max(visual.width * scale, visual.kind === 'line' ? 1 : 0);
  const height = Math.max(visual.height * scale, visual.kind === 'line' ? 1 : 0);
  const selected = field !== null && slot === activeSlot && field === activeField;
  const text = field === 'code' ? (item?.code || visual.text) : field === 'name' ? (item?.name || visual.text) : visual.text;
  const baseStyle = {
    position: 'absolute' as const,
    left,
    top,
    width,
    height,
    overflow: 'hidden' as const,
  };

  if (visual.kind === 'image' && visual.imageBase64) {
    return <Image
      key={visual.id}
      source={{ uri: `data:${visual.imageMime || 'image/png'};base64,${visual.imageBase64}` }}
      resizeMode="contain"
      style={baseStyle}
    />;
  }

  if (visual.kind === 'line') {
    const horizontal = visual.width >= visual.height;
    const thickness = Math.max(1, (visual.strokeWidth || 0.7) * 12700 * scale);
    return <View key={visual.id} pointerEvents="none" style={[
      baseStyle,
      horizontal
        ? { height: thickness, backgroundColor: visual.stroke || '#171717' }
        : { width: thickness, backgroundColor: visual.stroke || '#171717' },
    ]} />;
  }

  if (visual.kind === 'rect') {
    return <View key={visual.id} pointerEvents="none" style={[
      baseStyle,
      visual.fill ? { backgroundColor: visual.fill } : null,
      visual.stroke ? { borderWidth: Math.max(0.6, (visual.strokeWidth || 0.7) * 12700 * scale), borderColor: visual.stroke } : null,
    ]} />;
  }

  const fontSize = Math.max(4, Math.min(20, (visual.fontSizePt || 9) * 12700 * scale));
  const content = (
    <Text
      numberOfLines={field === 'name' ? 4 : 2}
      adjustsFontSizeToFit
      minimumFontScale={0.62}
      style={{
        color: visual.color || '#111111',
        fontSize,
        lineHeight: Math.max(fontSize * 1.05, 5),
        fontWeight: visual.bold ? '700' : '400',
        textAlign: visual.align || 'left',
        textAlignVertical: visual.verticalAlign === 'center' ? 'center' : visual.verticalAlign === 'bottom' ? 'bottom' : 'top',
        includeFontPadding: false,
      }}
    >
      {text}
    </Text>
  );

  if (field && slot >= 0) {
    return <Pressable
      key={visual.id}
      onPress={() => onSelectField(slot, field)}
      style={[baseStyle, selected ? { borderWidth: 1.5, borderColor: accent, backgroundColor: `${accent}18` } : null]}
    >
      {content}
    </Pressable>;
  }
  return <View key={visual.id} pointerEvents="none" style={baseStyle}>{content}</View>;
}

export default function TemplatePreview(props: Props) {
  const { width: screenWidth } = useWindowDimensions();
  const width = Math.max(220, Math.min(440, screenWidth - 72));
  const scale = width / Math.max(1, props.slide.widthEmu);
  const height = props.slide.heightEmu * scale;
  const visuals = [...(props.slide.visuals || [])].sort((a, b) => {
    // Text and field layers sit above borders and decorative geometry.
    const order = (value: TemplateVisual) => value.kind === 'rect' || value.kind === 'line' ? 0 : value.kind === 'image' ? 1 : 2;
    return order(a) - order(b);
  });

  return (
    <View style={{ alignItems: 'center', width: '100%' }}>
      <View style={{ width, height, backgroundColor: '#FFFFFF', overflow: 'hidden', borderRadius: 3, elevation: 3, shadowColor: '#101828', shadowOpacity: 0.14, shadowRadius: 9, shadowOffset: { width: 0, height: 4 } }}>
        {visuals.map(visual => <PreviewShape key={visual.id} {...props} visual={visual} scale={scale} />)}
      </View>
    </View>
  );
}
