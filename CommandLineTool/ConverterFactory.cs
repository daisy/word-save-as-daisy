using Daisy.SaveAsDAISY.Conversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.CommandLineTool {
    class ConverterFactory {
        private static WordToDTBookXMLTransform wordInstance;

        protected ConverterFactory() {
        }

        public static WordToDTBookXMLTransform Instance(Direction transformDirection) {
            switch (transformDirection) {
                case Direction.DocxToXml:
                    if (wordInstance == null) {
                        wordInstance = new WordToDTBookXMLTransform();
                    }
                    return wordInstance;
                default:
                    throw new ArgumentException("invalid transform direction type");
            }
        }
    }
}
