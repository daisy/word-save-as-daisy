using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion
{

    public class Alive
    {
        public bool alive;
        public bool localfs;
        public bool authentication;
        public string version;

        private Alive(bool alive, bool localfs, bool authentication, string version)
        {
            this.alive = alive;
            this.localfs = localfs;
            this.authentication = authentication;
            this.version = version;
        }

        public static Alive From(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNode root = doc.DocumentElement;
            if(root == null || root.Name != "alive")
                throw new Exception($"Invalid alive XML :\r\n{xml}");
            return new Alive(
                true,
                root.Attributes["localfs"]?.Value == "true",
                root.Attributes["authentication"]?.Value == "true",
                root.Attributes["version"]?.Value ?? "unknown"
            );
        }
    }

    public class ResultFile {
        public string MimeType;
        public string Size;
        public string File;
        public string Href;
    }

    public class NamedResult {
        public string From;
        public string Href;
        public string MimeType;
        public string Name;
        public string Nicename;
        public string Desc;
        public List<ResultFile> Files;
    }

    public class Results {
        public string Href;
        public string MimeType;
        public List<NamedResult> NamedResults;
    }

    public class Message {
        public string Level;
        public string Content;
        public int Sequence;
        public DateTime Timestamp;
        public List<Message> Messages = new List<Message>();
        
        public static Message From(XmlElement messageNode)
        {
            DateTime _ts = new DateTime();
            string timestamp = messageNode.GetAttribute("timeStamp");
            try {
                _ts = new DateTime(long.Parse(timestamp));
            }
            catch (Exception e) {
                Console.WriteLine(e);
            }

            Message message = new Message
            {
                Level = messageNode.GetAttribute("level"),
                Content = messageNode.GetAttribute("content"),
                Sequence = int.Parse(messageNode.GetAttribute("sequence")),
                Timestamp = _ts,
                Messages = new List<Message>()
            };

            foreach (XmlElement childNode in messageNode.ChildNodes) {
                message.Messages.Add(From(childNode));
            }

            return message;
        }
        public override string ToString()
        {
            return $"{Timestamp.ToString("HH:mm:ss.fff")} : {Content}";
        }

        /// <summary>
        /// Flatten the messages tree into a single list
        /// </summary>
        /// <returns></returns>
        public List<Message> Flatten()
        {
            List<Message> _messages = new List<Message>();
            _messages.Add(this);
            foreach (Message m in Messages) {
                _messages.AddRange(m.Flatten());
            }
            return _messages;
        }

    }

    public class JobData
    {
        public string _Type = "JobRequestSuccess"; // returned by the pipeline on success
        public string JobId; // the ID from the pipeline
        public string Priority;
        public string Status;
        public string Log;
        public Results Results;
        public string DownloadedFolder;
        public List<Message> Messages = new List<Message>();
        public double Progress;
        public ScriptData Script;
        public string Nicename;
        public string ScriptHref;
        public string Href;

        public DateTime lastMessageTimestamp = new DateTime();

        public string GetNewMessages(string prefix = "")
        {
            if(Messages.Count == 0)
                return string.Empty;
            // Flatten message list and keep only the new messages
            List<Message> messages = Messages.Aggregate(
                    new List<Message>(),
                    (res, next) => {
                        res.AddRange(next.Flatten());
                        return res;
                    },
                    (res) => res
                )
                .Where(m => m.Timestamp > lastMessageTimestamp)
                .OrderBy(m => m.Timestamp)
                .ToList();

            if (messages.Count == 0)
                return string.Empty;

            lastMessageTimestamp = messages.Last().Timestamp;
            return prefix + string.Join($"\r\n{prefix}",
                messages.Select(m => m.ToString())
            );
        }

        public static JobData[] FromJobs(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNode root = doc.DocumentElement;
            if (root == null || root.Name != "jobs")
                throw new Exception($"Invalid jobs list XML :\r\n{xml}");
            return root.SelectNodes("./*[local-name()='job']").Cast<XmlNode>().Select(
                (node) => FromJob(node.OuterXml)
            ).ToArray();
        }

        public static JobData FromJobRequest(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            // Check if root node is of type "error"
            if (doc.DocumentElement.Name == "error") {
                throw new Exception($"Pipeline returned an error : {doc.DocumentElement.InnerXml}");
            } else if (doc.DocumentElement.Name == "job") {
                return FromJob(doc.DocumentElement);
            } else {
                throw new Exception($"Invalid job response XML :\r\n{xml}");
            }
        }
        public static JobData FromJob(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlElement jobElm = doc.DocumentElement;
            if (jobElm == null || jobElm.Name != "job")
                throw new Exception($"Invalid job XML :\r\n{xml}");
            return FromJob(jobElm);

        }

        private static CultureInfo CI = (CultureInfo)CultureInfo.CurrentCulture.Clone();
       

        public static JobData FromJob(XmlElement jobElm)
        {
            JobData jobData = new JobData
            {
                JobId = jobElm.GetAttribute("id"),
                Priority = jobElm.GetAttribute("priority"),
                Status = jobElm.GetAttribute("status"),
                Href = jobElm.GetAttribute("href"),
                Nicename = jobElm.HasAttribute("nicename") ? jobElm.GetAttribute("nicename") : "Job"
            };

            XmlNodeList nicenameElms = jobElm.GetElementsByTagName("nicename");
            if (nicenameElms.Count > 0) {
                jobData.Nicename = nicenameElms[0].InnerText;
            }

            XmlNodeList logElms = jobElm.GetElementsByTagName("log");
            if (logElms.Count > 0) {
                jobData.Log = logElms[0].Attributes["href"]?.Value;
            }

            XmlNodeList resultsElms = jobElm.GetElementsByTagName("results");
            if (resultsElms.Count > 0) {
                Results results = new Results
                {
                    Href = resultsElms[0].Attributes["href"]?.Value,
                    MimeType = resultsElms[0].Attributes["mime-type"]?.Value,
                    NamedResults = new List<NamedResult>()
                };

                foreach (XmlElement resultElm in resultsElms[0].SelectNodes("./*[local-name()='result']")) {
                    if (resultElm.ParentNode == resultsElms[0]) {
                        NamedResult namedResult = new NamedResult
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
                            ResultFile resultFile = new ResultFile
                            {
                                MimeType = resultFileElm.GetAttribute("mime-type"),
                                Size = resultFileElm.GetAttribute("size"),
                                File = resultFileElm.GetAttribute("file"),
                                Href = resultFileElm.GetAttribute("href")
                            };
                            namedResult.Files.Add(resultFile);
                        }
                        results.NamedResults.Add(namedResult);
                    }
                }
                jobData.Results = results;
            }

            XmlNodeList messagesElms = jobElm.GetElementsByTagName("messages");
            if (messagesElms.Count > 0) {
                jobData.Messages = new List<Message>();
                foreach (XmlElement messageElm in messagesElms[0].ChildNodes) {
                    jobData.Messages.Add(Message.From(messageElm));
                }
                if(messagesElms[0].Attributes["progress"] != null && messagesElms[0].Attributes["progress"].Value.Length > 0) {
                    CI.NumberFormat.CurrencyDecimalSeparator = ".";
                    jobData.Progress = double.Parse(messagesElms[0].Attributes["progress"].Value, NumberStyles.Any, CI);
                }
                
            }

            XmlNodeList scriptElms = jobElm.GetElementsByTagName("script");
            if (scriptElms.Count > 0) {
                jobData.Script = ScriptData.FromScriptXML(scriptElms[0].OuterXml);
            }

            return jobData;
        }


    }


    public class ScriptItemBase{
        public string Desc;
        public string[] MediaType;
        public string Name;
        public bool Sequence;
        public bool Required;
        public string Nicename;
        public string Type;
        public string Kind;
        public bool Ordered;
        public string Pattern;
        public bool IsStylesheetParameter;
    }
    public class ScriptInput : ScriptItemBase {
        public string _Type = "anyFileURI";
        public string _Kind = "input";
        public new bool Ordered = false;

        public static ScriptInput From(XmlElement input)
        {
            return new ScriptInput()
            {
                Desc = input.Attributes["desc"]?.Value ?? "",
                MediaType = input.Attributes["mediaType"]?.Value.Split(' ') ?? new string[0],
                Name = input.Attributes["name"]?.Value.Trim() ?? "",
                Sequence = input.Attributes["sequence"]?.Value == "true",
                Required = input.Attributes["required"]?.Value == "true",
                Nicename = input.Attributes["nicename"]?.Value.Trim() ?? "",
                Ordered = false,
                Pattern = "",
                IsStylesheetParameter = false,
            };
        }
    }

    public class ScriptOption : ScriptItemBase {
        public string _Type;
        public string _Default;
        public string _Kind = "option";

        public static ScriptOption From(XmlElement input)
        {
            return new ScriptOption()
            {
                Desc = input.Attributes["desc"]?.Value ?? "",
                MediaType = input.Attributes["mediaType"]?.Value.Split(' ') ?? new string[0],
                Name = input.Attributes["name"]?.Value.Trim() ?? "",
                Sequence = input.Attributes["sequence"]?.Value == "true",
                Required = input.Attributes["required"]?.Value == "true",
                Nicename = input.Attributes["nicename"]?.Value.Trim() ?? "",
                Ordered = input.Attributes["ordered"]?.Value == "true",
                _Type = input.Attributes["type"]?.Value ?? "",
                _Default = input.Attributes["default"]?.Value ?? "",
                Pattern = "",
                IsStylesheetParameter = false,
            };
        }
    }

    public class ScriptData
    {
        public string Id;
        public string Href;
        public string Nicename;
        public string Description;
        public string Version;
        public string Homepage;
        public ScriptInput[] inputs;
        public ScriptOption[] options;

        private ScriptData() { }

        public static ScriptData[] fromScriptsXML(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlNode root = doc.DocumentElement;
            if (root == null || root.Name != "scripts")
                throw new Exception($"Invalid scripts list XML :\r\n{xml}");

            return root.SelectNodes("./*[local-name()='script']").Cast<XmlElement>().Select(
                (node) => FromScriptXML(node)
            ).ToArray();
        }
        
        public static ScriptData FromScriptXML(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlElement root = doc.DocumentElement;
            if (root == null || root.Name != "script")
                throw new Exception($"Invalid script XML :\r\n{xml}");
            return FromScriptXML(root);
        }

        public static ScriptData FromScriptXML(XmlElement root)
        {
            return new ScriptData()
            {
                Id = root.Attributes["id"]?.Value,
                Href = root.Attributes["href"]?.Value,
                Nicename = root.SelectSingleNode("./*[local-name()='nicename']")?.InnerText.Trim() ?? "",
                Description = root.SelectSingleNode("./*[local-name()='description']")?.InnerText.Trim() ?? "",
                Version = root.SelectSingleNode("./*[local-name()='version']")?.InnerText.Trim() ?? "",
                Homepage = root.SelectSingleNode("./*[local-name()='homepage']")?.InnerText.Trim() ?? "",
                inputs = root.SelectNodes("./*[local-name()='input']").Cast<XmlElement>().Select(
                    (input) => ScriptInput.From(input)
                ).ToArray(),
                options = root.SelectNodes("./*[local-name()='option']").Cast<XmlElement>().Select(
                    (input) => ScriptOption.From(input)
                ).ToArray(),
            };
        }

        public JobRequest makeRequest()
        {
            return new JobRequest()
            {
                Nicename = $"{Id} job",
                ScriptHref = Href,
                Inputs = inputs.Select(
                    input => new NameValue()
                    {
                        Name = input.Name,
                        Value = "",
                        IsFile = input._Type == "anyFileURI",
                        IsStylesheetParameter = input.IsStylesheetParameter
                    }
                ).ToList(),
                Options = options.Select(
                    option => new NameValue()
                    {
                        Name = option.Name,
                        Value = option._Default,
                        IsFile = option._Type == "anyFileURI",
                        IsStylesheetParameter = option.IsStylesheetParameter
                    }
                ).ToList(),

            };
        }

    }


    public class NameValue{
        public string Name;
        public object Value;
        public bool IsFile;
        public bool IsStylesheetParameter;
    }
    public class JobRequest
    {
        public string ScriptHref;
        public string Nicename = "Job";
        public string Priority = "medium";
        public string BatchId;
        public List<NameValue> Inputs;
        public List<NameValue> Options;
        public List<NameValue> StylesheetParameterOptions;
        public List<NameValue> Outputs;

        public static JobRequest fromScript(Webservice ws, Pipeline2Script s)
        {

            ScriptData matching = ws.FetchScripts().FirstOrDefault(script => script.Id == s.Name);
            if (matching == null) {
                throw new Exception($"Script {s.Name} not found in the pipeline");
            }
            JobRequest baseRequest = matching.makeRequest();
            
            foreach(var i in baseRequest.Inputs) {
                var kv = s.Parameters.FirstOrDefault(p => p.Value.Name == i.Name);
                if (kv.Value == null) {
                    continue;
                }
                i.Value = kv.Value.ParameterDataType is PathDataType && kv.Value.ParameterValue.ToString().Length > 0 ? new System.Uri(kv.Value.ParameterValue.ToString()).AbsoluteUri :
                    kv.Value.ParameterValue is string ? kv.Value.ParameterValue : kv.Value.ParameterValue.ToString().ToLower();
            }

            foreach (var i in baseRequest.Options) {
                var kv = s.Parameters.FirstOrDefault(p => p.Value.Name == i.Name);
                if(kv.Value == null) {
                    continue;
                }
                i.Value = kv.Value.ParameterDataType is PathDataType && kv.Value.ParameterValue.ToString().Length > 0 ? new System.Uri(kv.Value.ParameterValue.ToString()).AbsoluteUri :
                    kv.Value.ParameterValue is string ? kv.Value.ParameterValue : kv.Value.ParameterValue.ToString().ToLower();
            }
            return baseRequest;
        }

        public string ToXMLString()
        {
            return $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>
<jobRequest xmlns=""http://www.daisy.org/ns/pipeline/data"">
    <nicename>{Nicename}</nicename>
    <priority>{Priority}</priority>
    <script href=""{ScriptHref}""/>
    {string.Join("", Inputs
            .Where(input => input.Value != null && !string.IsNullOrWhiteSpace(input.Value.ToString()))
            .Select(input => $@"<input name=""{input.Name}"">{(input.Value is IEnumerable<object> inputValues
                    ? string.Join("", inputValues.Select(value => $@"<item value=""{value.ToString().Trim() ?? ""}""/>"))
                    : $@"<item value=""{input.Value.ToString().Trim() ?? ""}""/>")}</input>"))}
    {string.Join("", Options
            .Where(option => option.Value != null && !string.IsNullOrWhiteSpace(option.Value.ToString()))
            .Select(option => $@"<option name=""{option.Name}"">{(option.Value is IEnumerable<object> optionValues
                    ? string.Join("", optionValues.Select(value => $@"<item value=""{value.ToString().Trim() ?? ""}""/>"))
                    : option.Value.ToString().Trim() ?? "")}</option>"))}
</jobRequest>";
        }
        
    }

    /// <summary>
    /// DAISY Pipeline Webservice client wrapper
    /// </summary>
    public class Webservice
    {
        /// <summary>
        /// Host of the webservice
        /// </summary>
        public string Host { get; set; }

        /// <summary>
        /// Port of the webservice
        /// </summary>
        public string Port { get; set; }

        /// <summary>
        /// Base path of the Pipeline REST API
        /// </summary>
        public string BasePath { get; set; }

        /// <summary>
        /// SSL flag to use https
        /// </summary>
        public bool SSL { get; set; } = false;


        /// <summary>
        /// client used to connect to the pipeline webservice
        /// </summary>
        private HttpClient client = new HttpClient();


        /// <summary>
        /// URL of the webservice
        /// </summary>
        /// <returns></returns>
        private string BaseURL()
        {
            return (SSL ? "https" : "http") + $"://{this.Host}" + (this.Port.Length > 0 ? $":{this.Port}" : "") + this.BasePath;
        }

        /// <summary>
        /// Fetch the Alive status of the engine
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public Alive FetchAlive()
        {
            int maxTry = 3;
            int countTry = 0;
            while (countTry < maxTry) {
                countTry++;
                try {
                    return this.client.GetAsync($"{BaseURL()}/alive")
                        .ContinueWith((response) =>
                        {
                            if (!response.Result.IsSuccessStatusCode)
                                throw new Exception($"Failed to fetch alive status from {BaseURL()}/alive : {response.Result.StatusCode}");
                            return response.Result.Content.ReadAsStringAsync().Result;
                        })
                        .ContinueWith((text) => Alive.From(text.Result))
                        .Result;
                }
                catch (Exception e) {
                    Console.WriteLine($"Connexion error (pipeline could have been busy) retry {countTry} ...");
                    //Thread.Sleep(1000);
                }
            }
            throw new Exception("Failed to fetch alive status after 3 tries");
           
        }

        /// <summary>
        /// Pipeline scripts cache to avoid repetitive fetches
        /// </summary>
        private ScriptData[] scriptsCache = null;

        /// <summary>
        /// Retrieve the scripts available in the pipeline
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public ScriptData[] FetchScripts()
        {
            if(scriptsCache == null) {
                scriptsCache = this.client.GetAsync($"{BaseURL()}/scripts")
                .ContinueWith((response) => {
                    if (!response.Result.IsSuccessStatusCode)
                        throw new Exception($"Failed to fetch scripts from {BaseURL()}/scripts : {response.Result.StatusCode}");
                    return response.Result.Content.ReadAsStringAsync().Result;
                })
                .ContinueWith((text) => ScriptData.fromScriptsXML(text.Result))
                .ContinueWith((scripts) => scripts.Result.ToList().Select(s => FetchScriptDetails(s)).ToArray())
                .Result;
            }
            return scriptsCache;
        }

        /// <summary>
        /// Retrieve the full description of a script from the pipeline
        /// </summary>
        /// <param name="s">the script to get additionnal information from</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public ScriptData FetchScriptDetails(ScriptData s) {
            return this.client.GetAsync($"{s.Href}")
                .ContinueWith((response) => {
                    if (!response.Result.IsSuccessStatusCode)
                        throw new Exception($"Failed to fetch script details from {s.Href} : {response.Result.StatusCode}");
                    return response.Result.Content.ReadAsStringAsync().Result;
                })
                .ContinueWith((text) => ScriptData.FromScriptXML(text.Result))
                .Result;
        }

        /// <summary>
        /// Retrieve job current status and other details from the engine
        /// (The job has to exists in the engine)
        /// </summary>
        /// <param name="j"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public JobData FetchJobDetails(JobData j) {
            JobData result = this.client.GetAsync($"{j.Href}")
                .ContinueWith((response) => {
                    if (!response.Result.IsSuccessStatusCode)
                        throw new Exception($"Failed to fetch script details from {j.Href} : {response.Result.StatusCode}");
                    return response.Result.Content.ReadAsStringAsync().Result;
                })
                .ContinueWith((text) => JobData.FromJob(text.Result))
                .Result;
            result.lastMessageTimestamp = j.lastMessageTimestamp;
            return result;
        }

        /// <summary>
        /// Fetch the log of the job
        /// </summary>
        /// <param name="j"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public string FetchJobLog(JobData j) {
            return this.client.GetAsync($"{j.Log}")
                .ContinueWith((response) => {
                    if (!response.Result.IsSuccessStatusCode)
                        throw new Exception($"Failed to fetch script details from {j.Log} : {response.Result.StatusCode}");
                    return response.Result.Content.ReadAsStringAsync().Result;
                }).Result;
        }

        /// <summary>
        /// Launch a job on the engine
        /// </summary>
        /// <param name="j"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public JobData LaunchJob(JobRequest j) {
            string requestString = j.ToXMLString();
            return this.client.PostAsync(
                    $"{BaseURL()}/jobs",
                    new StringContent(requestString, Encoding.UTF8, "application/xml")
                )
                .ContinueWith((response) => {
                    if (!response.Result.IsSuccessStatusCode)
                        throw new Exception($"Failed to job request to {BaseURL()}/jobs : {response.Result.StatusCode}");
                    return response.Result.Content.ReadAsStringAsync().Result;
                })
                .ContinueWith((text) => JobData.FromJobRequest(text.Result))
                .Result;
        }

        /// <summary>
        /// Delete a job in the pipeline
        /// </summary>
        /// <param name="j">the job to delete</param>>
        /// <returns>the http response of the webserver. Expected codes are <br/>
        /// - 204 on job removal<br/>
        /// - 404 if the job does not exists (i.e. if it has already been removed)
        /// </returns>
        public HttpResponseMessage DeleteJob(JobData j ) {
            return this.client.DeleteAsync($"{j.Href}").Result;
            //.ContinueWith((response) => {
            //    if (!response.Result.IsSuccessStatusCode)
            //        throw new Exception($"Failed to delete job from {BaseURL()}/{j.Href} : {response.Result.StatusCode}");
            //}).Wait();
        }

        /// <summary>
        /// Rerieve the results (usually a zip file to be decompressed) as a byte array
        /// </summary>
        /// <param name="j"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public byte[] FetchResults(JobData j)
        {
            return this.client.GetAsync($"{j.Results.Href}")
                .ContinueWith((response) => {
                    if (!response.Result.IsSuccessStatusCode)
                        throw new Exception($"Failed to fetch script details from {j.Results.Href} : {response.Result.StatusCode}");
                    return response.Result.Content.ReadAsByteArrayAsync().Result;
                }).Result;
        }

       
    }
}
