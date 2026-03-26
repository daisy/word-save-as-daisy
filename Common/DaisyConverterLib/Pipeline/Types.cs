
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Types
{
    public class TtsVoice
    {
        public string Engine { get; set; }
        public string Name { get; set; }
        public string Lang { get; set; }
        public string Gender { get; set; }
        public int? Priority { get; set; }
        public string Id { get; set; }
        public string Href { get; set; }
        public string Preview { get; set; }
    }

    public class TtsEngineProperty
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public class TtsConfig
    {
        public List<TtsVoice> PreferredVoices { get; set; }
        public List<TtsVoice> DefaultVoices { get; set; }
        public List<TtsEngineProperty> TtsEngineProperties { get; set; }
        public string XmlFilepath { get; set; }
    }

    public class TtsEngineState
    {
        public string Status { get; set; } // Peut être "disabled", "connecting", "available", "disconnecting", "error"
        public List<string> Features { get; set; }
        public string Message { get; set; }
        public string Name { get; set; }
        public string VoicesUrl { get; set; }
    }

    public class WebserviceProperties
    {
        public string Host { get; set; }
        public int? Port { get; set; }
        public string Path { get; set; }
        public bool? Ssl { get; set; }
        public long? LastStart { get; set; }
    }

    public enum PipelineStatus
    {
        Unknown,
        Starting,
        Running,
        Stopped,
        Error
    }

    public class PipelineState
    {
        public WebserviceProperties Webservice { get; set; }
        public PipelineStatus Status { get; set; }
        public List<Job> Jobs { get; set; }
        public List<ScriptDefinition> Scripts { get; set; }
        public List<TtsVoice> TtsVoices { get; set; }
        public int? InternalJobCounter { get; set; }
        public string SelectedJobId { get; set; }
        public List<Datatype> Datatypes { get; set; }
        public AliveData Alive { get; set; }
        public Dictionary<string, EngineProperty> Properties { get; set; }
        public Dictionary<string, TtsEngineState> TtsEnginesStates { get; set; }
    }

    public enum PipelineType
    {
        Embedded
        // System,
        // Remote
    }

    public class PipelineInstanceProperties
    {
        public PipelineType? PipelineType { get; set; }
        public string PipelineHome { get; set; }
        public string AppDataFolder { get; set; }
        public string LogsFolder { get; set; }
        public string JrePath { get; set; }
        public WebserviceProperties Webservice { get; set; }
        // Les delegates pour onError/onMessage sont omis car ils ne sont pas utilisés de la même façon en C#
    }

    public class EngineProperty
    {
        public string Name { get; set; }
        public string Href { get; set; }
        public string Desc { get; set; }
        public string Value { get; set; }

        public static EngineProperty FromXml(XmlElement propertyElem)
        {
            return new EngineProperty
            {
                Name = propertyElem.GetAttribute("name"),
                Href = propertyElem.GetAttribute("href"),
                Desc = propertyElem.GetAttribute("desc"),
                Value = propertyElem.GetAttribute("value")
            };
        }

        public static List<EngineProperty> ListFromXml(XmlElement propertiesElem)
        {
            var properties = new List<EngineProperty>();
            foreach (XmlElement propertyElem in propertiesElem.GetElementsByTagName("property")) {
                // On ne garde que les enfants directs
                if (propertyElem.ParentNode == propertiesElem) {
                    properties.Add(FromXml(propertyElem));
                }
            }
            return properties;
        }

        public override string ToString()
        {
            return $"{Name} = {Value} ({Desc})";
        }

        public string ToUpdateString()
        {
            return $"<property  xmlns=\"http://www.daisy.org/ns/pipeline/data\" name=\"{Name}\" value=\"{(Value == null ? "" : Value)}\"/>";
        }

        public XmlElement ToXml(XmlDocument doc)
        {
            XmlElement propertyElem = doc.CreateElement("property");
            propertyElem.SetAttribute("name", Name);
            propertyElem.SetAttribute("href", Href);
            propertyElem.SetAttribute("desc", Desc);
            propertyElem.SetAttribute("value", Value);
            return propertyElem;
        }
    }

    public class AliveData
    {
        public bool Alive { get; set; }
        public bool? Localfs { get; set; }
        public bool? Authentication { get; set; }
        public string Version { get; set; }
    }

    public enum Priority
    {
        Low,
        Medium,
        High
    }

    public enum JobStatus
    {
        Success,
        Error,
        Idle,
        Running,
        Fail
    }

    public class ResultFile
    {
        public string MimeType { get; set; }
        public string Size { get; set; }
        public string File { get; set; }
        public string Href { get; set; }
    }

    public class NamedResult
    {
        public string From { get; set; }
        public string Href { get; set; }
        public string MimeType { get; set; }
        public string Name { get; set; }
        public string Nicename { get; set; }
        public string Desc { get; set; }
        public List<ResultFile> Files { get; set; }
    }

    public class Results
    {
        public string Href { get; set; }
        public string MimeType { get; set; }
        public List<NamedResult> NamedResults { get; set; }
    }

    public enum MessageLevel
    {
        Error,
        Warning,
        Info,
        Debug,
        Trace
    }

    public class Message
    {
        public MessageLevel Level { get; set; }
        public string Content { get; set; }
        public int Sequence { get; set; }
        public long Timestamp { get; set; }
        public List<Message> Messages { get; set; }
    }

    public enum JobState
    {
        New,
        Submitting,
        Submitted,
        Ended
    }

    public class Job
    {
        public string InternalId { get; set; }
        public int? Index { get; set; }
        public JobState State { get; set; }
        public JobData JobData { get; set; }
        public JobRequest JobRequest { get; set; }
        public JobRequestError JobRequestError { get; set; }
        public ScriptDefinition Script { get; set; }
        public List<JobError> Errors { get; set; }
        public string LinkedTo { get; set; }
        public bool? Invisible { get; set; }
        public List<ScriptOption> StylesheetParameters { get; set; }
        public bool? Is2StepsJob { get; set; }
        public bool? IsPrimaryForBatch { get; set; }
    }

    public class JobError
    {
        public string FieldName { get; set; }
        public string Error { get; set; }
    }

    public class JobData
    {
        public string Type { get; set; }
        public string JobId { get; set; }
        public Priority? Priority { get; set; }
        public JobStatus? Status { get; set; }
        public string Log { get; set; }
        public Results Results { get; set; }
        public string DownloadedFolder { get; set; }
        public List<Message> Messages { get; set; }
        public double? Progress { get; set; }
        public ScriptDefinition Script { get; set; }
        public string Nicename { get; set; }
        public string ScriptHref { get; set; }
        public string Href { get; set; }

        public static JobData FromXml(XmlElement jobElm)
        {

            var jobData = new JobData
            {
                JobId = jobElm.GetAttribute("id"),
                Priority = ParseEnumOrNull<Priority>(jobElm.GetAttribute("priority")),
                Status = ParseEnumOrNull<JobStatus>(jobElm.GetAttribute("status")),
                Href = jobElm.GetAttribute("href")
            };

            // nicename : élément ou attribut
            var nicenameElms = jobElm.GetElementsByTagName("nicename");
            if (nicenameElms.Count > 0)
                jobData.Nicename = nicenameElms[0].InnerText;
            else if (jobElm.HasAttribute("nicename"))
                jobData.Nicename = jobElm.GetAttribute("nicename");
            else
                jobData.Nicename = "Job";

            // log
            var logElms = jobElm.GetElementsByTagName("log");
            if (logElms.Count > 0) {
                jobData.Log = (logElms[0] as XmlElement).GetAttribute("href");
            }

            // results
            var resultsElms = jobElm.GetElementsByTagName("results");
            if (resultsElms.Count > 0) {
                var resultsElm = (XmlElement)resultsElms[0];
                var results = new Results
                {
                    Href = resultsElm.GetAttribute("href"),
                    MimeType = resultsElm.GetAttribute("mime-type"),
                    NamedResults = new List<NamedResult>()
                };

                foreach (XmlElement resultElm in resultsElm.GetElementsByTagName("result")) {
                    // On ne garde que les enfants directs
                    if (resultElm.ParentNode == resultsElm) {
                        var namedResult = new NamedResult
                        {
                            From = resultElm.GetAttribute("from"),
                            Href = resultElm.GetAttribute("href"),
                            MimeType = resultElm.GetAttribute("mime-type"),
                            Name = resultElm.GetAttribute("name"),
                            Nicename = resultElm.GetAttribute("nicename"),
                            Desc = resultElm.GetAttribute("desc"),
                            Files = new List<ResultFile>()
                        };

                        foreach (XmlElement resultFileElm in resultElm.GetElementsByTagName("result")) {
                            var resultFile = new ResultFile
                            {
                                MimeType = resultFileElm.GetAttribute("mime-type"),
                                Size = resultFileElm.GetAttribute("size"),
                                File = resultFileElm.HasAttribute("file") ? resultFileElm.GetAttribute("file") : null,
                                Href = resultFileElm.HasAttribute("href") ? resultFileElm.GetAttribute("href") : null
                            };
                            namedResult.Files.Add(resultFile);
                        }
                        results.NamedResults.Add(namedResult);
                    }
                }
                jobData.Results = results;
            }

            // messages
            var messagesElms = jobElm.GetElementsByTagName("messages");
            if (messagesElms.Count > 0) {
                var messagesElm = (XmlElement)messagesElms[0];
                try {
                    jobData.Messages = ArrayFromChildNodes(messagesElm.ChildNodes)
                        .Select(e => ParseMessage(e))
                        .ToList();
                }
                catch (Exception e) {
                    // Gérer l'erreur de parsing si besoin
                }
                var progressAttr = messagesElm.GetAttribute("progress");
                if (!string.IsNullOrEmpty(progressAttr) && double.TryParse(progressAttr, NumberStyles.Any, CultureInfo.InvariantCulture, out double progress))
                    jobData.Progress = progress;
            }

            // script
            var scriptElms = jobElm.GetElementsByTagName("script");
            if (scriptElms.Count > 0) {
                jobData.Script = ScriptDefinition.FromXml((XmlElement)scriptElms[0]);
            }

            return jobData;
        }

        private static List<XmlElement> ArrayFromChildNodes(XmlNodeList nodeList)
        {
            var result = new List<XmlElement>();
            foreach (XmlNode node in nodeList) {
                if (node is XmlElement element)
                    result.Add(element);
            }
            return result;
        }

        private static Message ParseMessage(XmlElement messageNode)
        {
            long timestamp = 0;
            var timestampAttr = messageNode.GetAttribute("timeStamp");
            if (!string.IsNullOrEmpty(timestampAttr)) {
                long.TryParse(timestampAttr, out timestamp);
            }

            var message = new Message
            {
                Level = ParseEnumOrNull<MessageLevel>(messageNode.GetAttribute("level")) ?? MessageLevel.Info,
                Content = messageNode.GetAttribute("content"),
                Sequence = int.TryParse(messageNode.GetAttribute("sequence"), out int seq) ? seq : 0,
                Timestamp = timestamp,
                Messages = ArrayFromChildNodes(messageNode.ChildNodes).Select(e => ParseMessage(e)).ToList()
            };
            return message;
        }

        private static T? ParseEnumOrNull<T>(string value) where T : struct
        {
            if (string.IsNullOrEmpty(value)) return null;
            try {
                return (T)Enum.Parse(typeof(T), value, true);
            }
            catch {
                return null;
            }
        }
    }

    public class JobRequestError : Exception
    {
        public string Type { get; set; }
        public string Description { get; set; }
        public string Trace { get; set; }

        public override string ToString()
        {
            return $"{Type} - {Description}\nTrace: {Trace}";
        }
    }

    public class ScriptItemBase
    {
        public string Desc { get; set; }
        public List<string> MediaType { get; set; }
        public string Name { get; set; }
        public bool? Sequence { get; set; }
        public bool? Required { get; set; }
        public string Nicename { get; set; }
        public string Type { get; set; }
        public string Kind { get; set; }
        public bool? Ordered { get; set; }
        public string Pattern { get; set; }
        public bool? IsStylesheetParameter { get; set; }
    }

    public class ScriptInput : ScriptItemBase
    {
        public new string Type { get; set; } = "anyFileURI";
        public new string Kind { get; set; } = "input";
        public new bool Ordered { get; set; }
        public bool Batchable { get; set; }

        public static ScriptInput FromXml(XmlElement inputElm)
        {
            var mediaType = new List<string>();
            string mediaTypeVal = inputElm.GetAttribute("mediaType") ?? "";
            if (mediaTypeVal != "") {
                mediaType = mediaTypeVal.Split(' ').ToList();
            }
            return new ScriptInput()
            {
                Desc = inputElm.GetAttribute("desc") ?? "",
                MediaType = mediaType,
                Name = inputElm.GetAttribute("name") ?? "",
                Sequence = inputElm.GetAttribute("sequence") == "true",
                Required = inputElm.GetAttribute("required") == "true",
                Nicename = inputElm.GetAttribute("nicename") ?? "",
                Type = "anyFileURI",
                Kind = "input",
                Ordered = false,
                IsStylesheetParameter = false,
                Batchable = false,
            };
        }
    }
    public class ScriptOutput : ScriptItemBase
    {
        public new string Type { get; set; } = "anyDirURI";
        public new string Kind { get; set; } = "output";
        public new bool Ordered { get; set; }
        public bool Batchable { get; set; }

        public static ScriptOutput FromXml(XmlElement outputElm)
        {
            var mediaType = new List<string>();
            string mediaTypeVal = outputElm.GetAttribute("mediaType") ?? "";
            if (mediaTypeVal != "")
            {
                mediaType = mediaTypeVal.Split(' ').ToList();
            }
            return new ScriptOutput()
            {
                Desc = outputElm.GetAttribute("desc") ?? "",
                MediaType = mediaType,
                Name = outputElm.GetAttribute("name") ?? "",
                Sequence = outputElm.GetAttribute("sequence") == "true",
                Required = outputElm.GetAttribute("required") == "true",
                Nicename = outputElm.GetAttribute("nicename") ?? "",
                Type = "anyFileURI",
                Kind = "output",
                Ordered = false,
                IsStylesheetParameter = false,
                Batchable = false,
            };
        }
    }

    public class ScriptOption : ScriptItemBase
    {
        public new string Kind { get; set; } = "option";
        public string Default { get; set; }

        public static ScriptOption FromXml(XmlElement optionElm)
        {
            var mediaType = new List<string>();
            string mediaTypeVal = optionElm.GetAttribute("mediaType") ?? "";
            if (mediaTypeVal != "") {
                mediaType = mediaTypeVal.Split(' ').ToList();
            }
            return new ScriptOption()
            {
                Desc = optionElm.GetAttribute("desc") ?? "",
                MediaType = mediaType,
                Name = optionElm.GetAttribute("name") ?? "",
                Sequence = optionElm.GetAttribute("sequence") == "true",
                Required = optionElm.GetAttribute("required") == "true",
                Nicename = optionElm.GetAttribute("nicename"),
                Ordered = optionElm.GetAttribute("ordered") == "true",
                Type = optionElm.GetAttribute("type"),
                Default = optionElm.GetAttribute("default") ?? "",
                Kind = "option",
                IsStylesheetParameter = false,
            };
        }
    }

    public class ScriptDefinition
    {
        public string Id { get; set; }
        public string Href { get; set; }
        public string Nicename { get; set; }
        public string Description { get; set; }
        public string Version { get; set; }
        public List<ScriptInput> Inputs { get; set; }
        public List<ScriptOption> Options { get; set; }

        public List<ScriptOutput> Outputs { get; set; }

        public string Homepage { get; set; }
        public bool Batchable { get; set; }
        public bool Multidoc { get; set; }

        public static List<ScriptDefinition> ListFromXml(XmlElement scriptsElem)
        {
            var scripts = new List<ScriptDefinition>();
            foreach (XmlElement scriptElm in scriptsElem.GetElementsByTagName("script")) {
                // On ne garde que les enfants directs
                if (scriptElm.ParentNode == scriptsElem) {
                    scripts.Add(FromXml(scriptElm));
                }
            }
            return scripts;
        }

        public static ScriptDefinition FromXml(XmlElement scriptElm)
        {
            var inputs = new List<ScriptInput>();
            var options = new List<ScriptOption>();
            var outputs = new List<ScriptOutput>();
            ScriptInput firstRequired = null;
            foreach (XmlElement inputElm in scriptElm.GetElementsByTagName("input")) {
                ScriptInput scriptInput = ScriptInput.FromXml(inputElm);
                if(scriptInput.Required == true && firstRequired == null) {
                    firstRequired = scriptInput;
                }
                inputs.Add(scriptInput);
            }
            foreach (XmlElement optionElm in scriptElm.GetElementsByTagName("option")) {
                options.Add(ScriptOption.FromXml(optionElm));
            }

            foreach (XmlElement outputElm in scriptElm.GetElementsByTagName("output"))
            {
                outputs.Add(ScriptOutput.FromXml(outputElm));
            }


            return new ScriptDefinition()
            {
                Id = scriptElm.GetAttribute("id") ?? "",
                Href = scriptElm.GetAttribute("href") ?? "",
                Nicename = scriptElm.GetElementsByTagName("nicename")?.Item(0)?.InnerText ?? "",
                Description = scriptElm.GetElementsByTagName("description")?.Item(0)?.InnerText ?? "",
                Version = scriptElm.GetElementsByTagName("version")?.Item(0)?.InnerText ?? "",
                Inputs = inputs,
                Outputs = outputs,
                Options = options,
                Homepage = scriptElm.GetElementsByTagName("homepage")?.Item(0)?.InnerText ?? "",
                Batchable = options.Find(o => o.Name == "stylesheet-parameters") == null && inputs.Find(input => input.Sequence == true) == null,
                Multidoc = firstRequired?.Sequence == true,
            };
        }

    }

    public class NameValue
    {
        public string Name { get; set; }
        public object Value { get; set; }
        public string Type { get; set; }
        public bool IsStylesheetParameter { get; set; }
    }

    public class Callback
    {
        public string Href { get; set; }
        public List<string> Type { get; set; }
        public string Frequency { get; set; }
    }

    public class JobRequest
    {
        public string ScriptHref { get; set; }
        public string Nicename { get; set; }
        public List<string> Priority { get; set; }
        public string BatchId { get; set; }
        public List<NameValue> Inputs { get; set; }
        public List<NameValue> Options { get; set; }
        public List<NameValue> StylesheetParameterOptions { get; set; }
        public List<NameValue> Outputs { get; set; }
        public List<Callback> Callbacks { get; set; }

        public string ToXmlString()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>");
            stringBuilder.Append("<jobRequest xmlns=\"http://www.daisy.org/ns/pipeline/data\">");
            stringBuilder.Append($"<nicename>{Nicename}</nicename>");
            stringBuilder.Append("<priority>medium</priority>");
            stringBuilder.Append($"<script href=\"{ScriptHref}\"/>");
            foreach(var input in Inputs.Where(i => i.Value != null && i.Value.ToString().Trim() != "")) {
                stringBuilder.Append($"<input name=\"{input.Name}\">");
                if(input.Value is string) {
                    stringBuilder.Append($"<item value=\"{input.Value.ToString().Trim()}\"/>");
                } else if (input.Value is List<string>) {
                    foreach (var value in (List<string>)input.Value) {
                        stringBuilder.Append($"<item value=\"{value.ToString().Trim()}\"/>");
                    }
                }
                stringBuilder.Append("</input>");
            }
            foreach (var option in Options.Where(o => o.Value != null && o.Value.ToString().Trim() != "")) {
                stringBuilder.Append($"<option name=\"{option.Name}\">");
                if (option.Value is bool) {
                    stringBuilder.Append(option.Value.ToString().ToLower().Trim());
                } else if (option.Value is string) {
                    stringBuilder.Append(option.Value.ToString().Trim());
                } else if (option.Value is List<string>) {
                    foreach (var value in (List<string>)option.Value) {
                        stringBuilder.Append($"<item value=\"{value.ToString().Trim()}\"/>");
                    }
                } else {
                    stringBuilder.Append(option.Value.ToString().Trim());
                }
                stringBuilder.Append("</option>");
            }
            stringBuilder.Append("</jobRequest>");
            return stringBuilder.ToString();
        }
    }

    public class Datatype
    {
        public string Href { get; set; }
        public string Id { get; set; }

        // TODO : datatypes can be either of type "choice" with enum of <value></value><documentation></documentation>
        // or of type "data" with type string that contains a element a:documentation and an element "param" @type=pattern
        public List<DatatypeChoice> Choices { get; set; }

        public static Datatype FromXml(XmlElement datatypeElm)
        {
           // T
            return new Datatype
            {
                Href = datatypeElm.GetAttribute("href"),
                Id = datatypeElm.GetAttribute("id"),
                
            };
        }
    }

    

    public class DatatypeChoice
    {
        public string Documentation { get; set; }

        //public static DatatypeChoice FromXml(XmlElement choiceElm)
        //{
            
   //         var children =  Array.from(choiceElm.childNodes).filter(
   //    (n) => n.nodeType == n.ELEMENT_NODE
   //)
   // let choices = []
   // for (let i = 0; i < children.length; i++)
   //         {
   //             let c = children[i]
   //     //@ts-ignore
   //     if (c.tagName == 'data')
   //             {
   //                 choices.push(dataElementToJson(c))
   //         //@ts-ignore
   //     }
   //             else if (c.tagName == 'value')
   //             {
   //                 //@ts-ignore
   //                 let value = c.textContent?.trim()
   //         let documentation = ''
   //         // look ahead for a sibling 'documentation' element
   //         if (i + 1 < children.length)
   //                 {
   //                     if (
   //                         //@ts-ignore
   //                         children[i + 1].tagName == 'documentation' ||
   //                         //@ts-ignore
   //                         children[i + 1].tagName == 'a:documentation'
   //                     )
   //                     {
   //                         //@ts-ignore
   //                         documentation = children[++i].textContent.trim()
   //                     }
   //                 }
   //                 let valueChoice: ValueChoice = { value, documentation }
   //                 choices.push(valueChoice)
   //     }
   //         }
   //         return choices
        //}
    }

    public class ValueChoice : DatatypeChoice
    {
        public string Value { get; set; }
    }

    public class TypeChoice : DatatypeChoice
    {
        public string Type { get; set; }
        public string Pattern { get; set; }

        public static TypeChoice FromXml(XmlElement datatypeElm)
        {
            var documentationElms = datatypeElm.GetElementsByTagName("documentation");
            var paramElms = datatypeElm.GetElementsByTagName("param");
            var retval = new TypeChoice
            {
                Type = datatypeElm.GetAttribute("type"),
            };
            if (documentationElms.Count > 0)
            {
                retval.Documentation = documentationElms[0].InnerText.Trim();
            }
            if (
                paramElms.Count > 0)
            {
                XmlElement param = (XmlElement)paramElms[0];
                if (param.HasAttribute("name") && param.GetAttribute("name") == "pattern")
                {
                    retval.Pattern = param.InnerText.Trim();
                }
            }
            return retval;
        }
    }

    /*
     * 
    function datatypeXmlToJson(href, id, xmlString): Datatype {
    let root = sniffRoot(xmlString)
    let datatype = { id, href }
    if (root == 'choice') {
        let choiceElm = parseXml(xmlString, 'choice')
        return { ...datatype, choices: choiceElementToJson(choiceElm) }
    } else if (root == 'data') {
        let dataElm = parseXml(xmlString, 'data')
        return { ...datatype, choices: [dataElementToJson(dataElm)] }
    } else {
        return null
    }
}

function choiceElementToJson(choiceElm: Element): DatatypeChoice[] {
    //@ts-ignore
    let children = Array.from(choiceElm.childNodes).filter(
        (n) => n.nodeType == n.ELEMENT_NODE
    )
    let choices = []
    for (let i = 0; i < children.length; i++) {
        let c = children[i]
        //@ts-ignore
        if (c.tagName == 'data') {
            choices.push(dataElementToJson(c))
            //@ts-ignore
        } else if (c.tagName == 'value') {
            //@ts-ignore
            let value = c.textContent?.trim()
            let documentation = ''
            // look ahead for a sibling 'documentation' element
            if (i + 1 < children.length) {
                if (
                    //@ts-ignore
                    children[i + 1].tagName == 'documentation' ||
                    //@ts-ignore
                    children[i + 1].tagName == 'a:documentation'
                ) {
                    //@ts-ignore
                    documentation = children[++i].textContent.trim()
                }
            }
            let valueChoice: ValueChoice = { value, documentation }
            choices.push(valueChoice)
        }
    }
    return choices
}

function dataElementToJson(dataElm): TypeChoice {
    let documentationElms = dataElm.getElementsByTagName('documentation')
    let paramElms = dataElm.getElementsByTagName('param')
    let retval: TypeChoice = {
        type: dataElm.getAttribute('type'),
    }
    if (documentationElms.length > 0) {
        retval.documentation = documentationElms[0].textContent.trim()
    }
    if (
        paramElms.length > 0 &&
        paramElms[0].hasAttribute('name') &&
        paramElms[0].getAttribute('name') == 'pattern'
    ) {
        retval.pattern = paramElms[0].textContent.trim()
    }
    return retval
}
    */

    public class Filetype
    {
        public string Name { get; set; }
        public List<string> Extensions { get; set; }
        public string Type { get; set; }
    }
}