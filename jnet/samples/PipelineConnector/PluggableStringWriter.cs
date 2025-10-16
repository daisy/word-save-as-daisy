using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipelineConnector {
    /// <summary>
    /// String writer extension that sends back event when called
    /// https://stackoverflow.com/questions/3708454/is-there-a-textwriter-child-class-that-fires-event-if-text-is-written
    /// </summary>
    public class PluggableStringWriter : StringWriter {
        [EditorBrowsable(EditorBrowsableState.Never)]
        public delegate void FlushedEventHandler(object sender, EventArgs args);

        public event FlushedEventHandler Flushed;

        public FlushedEventHandler OnFlushEvent { get => Flushed; set => Flushed = value; }

        public virtual bool AutoFlush { get; set; }

        public PluggableStringWriter()
            : base() { }

        public PluggableStringWriter(bool autoFlush)
            : base() { this.AutoFlush = autoFlush; }

        protected void OnFlush() {
            var eh = Flushed;
            if (eh != null)
                eh(this, EventArgs.Empty);
        }

        public override void Flush() {
            base.Flush();
            OnFlush();
        }

        public override void Write(char value) {
            base.Write(value);
            if (AutoFlush) Flush();
        }

        public override void Write(string value) {
            base.Write(value);
            if (AutoFlush) Flush();
        }

        public override void Write(char[] buffer, int index, int count) {
            base.Write(buffer, index, count);
            if (AutoFlush) Flush();
        }
    }
}
