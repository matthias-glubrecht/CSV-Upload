// TextDecoder is available in all modern browsers but missing from the
// TypeScript DOM lib shipped with SPFx 1.4.1 (TypeScript ~2.4).

// tslint:disable-next-line:interface-name
interface TextDecoderOptions {
  fatal?: boolean;
  ignoreBOM?: boolean;
}

// tslint:disable-next-line:interface-name
interface TextDecoder {
  readonly encoding: string;
  readonly fatal: boolean;
  readonly ignoreBOM: boolean;
  decode(input?: ArrayBuffer | ArrayBufferView, options?: { stream?: boolean }): string;
}

// tslint:disable-next-line:no-shadowed-variable
declare var TextDecoder: {
  prototype: TextDecoder;
  new(label?: string, options?: TextDecoderOptions): TextDecoder;
};
