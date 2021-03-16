using Daisy.DaisyConverter.DaisyConverterLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.DaisyConverter.CommandLineTool {
    class ConverterFactory {
        private static AbstractConverter wordInstance;

        protected ConverterFactory() {
        }

        public static AbstractConverter Instance(Direction transformDirection) {
            switch (transformDirection) {
                case Direction.DocxToXml:
                    if (wordInstance == null) {
                        wordInstance = new Daisy.DaisyConverter.Word.Converter();
                    }
                    return wordInstance;
                default:
                    throw new ArgumentException("invalid transform direction type");
            }
        }
    }
}
