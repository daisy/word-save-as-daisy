using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    public class Webservice
    {
        public string Host { get; private set; }
        public int Port { get; private set; }

        public string Path { get; private set; }

        private HttpClient connection;

        private List<ScriptDefinition> scripts = null;

        internal static class ENDPOINTS
        {
            public static readonly string SCRIPTS = "/scripts";
            public static readonly string JOBS = "/jobs";
            public static readonly string DATATYPES = "/datatypes";
            public static readonly string ALIVE = "/alive";
            public static readonly string VOICES = "/voices";
            public static readonly string ADMIN_PROPERTIES = "/admin/properties";
            public static readonly string ADMIN_PROPERTY = "/admin/properties/{0}";
            public static readonly string TTS_ENGINES = "/tts-engines";
            public static readonly string STYLESHEET_PARAMETERS = "/stylesheet-parameters";
        }

        public Webservice(string host = "127.0.0.1", int port = 8080, string path = "/ws")
        {
            Host = host;
            Port = port;
            Path = path.StartsWith("/") ? path.TrimEnd('/') : $"/{path.TrimEnd('/')}";

            connection = new HttpClient
            {
                BaseAddress = new Uri($"http://{Host}:{Port}{Path}")
            };
        }

        private static XmlElement ParseXml(string xmlString, string rootElementName)
        {
            var doc = new XmlDocument();
            doc.LoadXml(xmlString);
            var root = doc.DocumentElement;
            if (root != null && root.Name == rootElementName)
                return root;
            else throw new Exception($"Invalid response: expected root element '{rootElementName}' not found.");
        }

        public async Task<AliveData> Alive()
        {
            try {
                var response = await connection.GetAsync(connection.BaseAddress + ENDPOINTS.ALIVE);
                if (!response.IsSuccessStatusCode) {
                    return new AliveData() { Alive = false };
                }


                XmlDocument doc = new XmlDocument();
                doc.LoadXml(await response.Content.ReadAsStringAsync());
                XmlElement aliveElement = ParseXml(await response.Content.ReadAsStringAsync(), "alive");
                return new AliveData()
                {
                    Alive = true,
                    Localfs = aliveElement.GetAttribute("localfs") == "true",
                    Authentication = aliveElement.GetAttribute("authentication") == "true",
                    Version = aliveElement.GetAttribute("version")
                };

            }
            catch (Exception e) {
                AddinLogger.Error(new Exception($"Error checking webservice alive status", e));
                return new AliveData() { Alive = false };
            }
        }

        public async Task<List<ScriptDefinition>> GetScripts()
        {
            if (scripts != null) return scripts;
            try {
                var response = await connection.GetAsync(connection.BaseAddress + ENDPOINTS.SCRIPTS);
                scripts = new List<ScriptDefinition>();
                XmlElement scriptsElm = ParseXml(await response.Content.ReadAsStringAsync(), "scripts");

                foreach (XmlElement scriptNode in scriptsElm.GetElementsByTagName("script")) {
                    scripts.Add(ScriptDefinition.FromXml(scriptNode));
                }
                scripts = Task.WhenAll(scripts.Select(s => GetScriptDetails(s))).Result.ToList();
                
                return scripts;

            }
            catch (Exception e) {
                scripts = null;
                AddinLogger.Error(new Exception($"Could not retrieve the list of scripts from the engine", e));
                return new List<ScriptDefinition>();
            }
        }
        
        public async Task<ScriptDefinition> GetScriptDetails(ScriptDefinition s)
        {   
            try {
                var response = await connection.GetAsync(s.Href);
                return ScriptDefinition.FromXml(ParseXml(await response.Content.ReadAsStringAsync(), "script"));
            }
            catch (Exception e) {
                AddinLogger.Error(new Exception($"Could not retrieve the details of script {s.Id} from the engine", e));
                return s;
            }
        }


        public async Task<JobData> LaunchJob(string scriptName, Dictionary<string, object> options = null)
        {
            try {
                var _scripts = await GetScripts();
                if (_scripts == null || !_scripts.Any()) {
                    throw new JobRequestError
                    {
                        Type = "JobUnknownResponse",
                        Description = "No scripts available on the server.",
                        Trace = "",
                    };
                }
                var script = _scripts.FirstOrDefault(s => s.Id == scriptName);
                if (script == null) {
                    throw new JobRequestError
                    {
                        Type = "JobRequestError",
                        Description = $"Script '{scriptName}' not found on the server.",
                        Trace = "",
                    };
                }
                var jobRequest = new Types.JobRequest
                {
                    ScriptHref = script.Href,
                    Nicename = "",
                    Inputs = new List<NameValue>(),
                    Options = new List<NameValue>(),
                };
                if (options != null) {
                    List<string> inputsNames = script.Inputs.Select(p => p.Name).ToList();
                    List<string> optionsNames = script.Options.Select(p => p.Name).ToList();
                    foreach (var option in options) {
                        if (inputsNames.Contains(option.Key)) {
                            jobRequest.Inputs.Add(new NameValue()
                            {
                                Name = option.Key,
                                Value = option.Value?.ToString() ?? ""
                            });
                        } else if (optionsNames.Contains(option.Key)) {
                            jobRequest.Options.Add(new NameValue()
                            {
                                Name = option.Key,
                                Value = option.Value?.ToString() ?? ""
                            });
                        } else {
                            //throw new Exception($"Option '{option.Key}' is not a valid input or option for script '{scriptName}'.");
                        }
                    }
                }
                return await LaunchJob(jobRequest);
            } catch (Exception e) {
                AddinLogger.Error(new Exception($"Error launching job for script '{scriptName}'", e));
                throw e;
            }

        }

        public async Task<JobData> LaunchJob(Types.JobRequest jobRequest)
        {
            try {
                var response = await connection.PostAsync(connection.BaseAddress + ENDPOINTS.JOBS, 
                    new StringContent(jobRequest.ToXmlString(), System.Text.Encoding.UTF8, "application/xml"));

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(await response.Content.ReadAsStringAsync());
                if(doc.DocumentElement == null) {
                    throw new Exception("Invalid response: expected root element 'job' not found.");
                }
                if(doc.DocumentElement.Name == "error") {
                    var descElms = doc.DocumentElement.GetElementsByTagName("description");
                    var traceElms = doc.DocumentElement.GetElementsByTagName("trace");
                    throw new JobRequestError
                    {
                        Type = "JobRequestError",
                        Description = descElms.Count > 0 ? descElms[0].InnerText.Trim() : "",
                        Trace = traceElms.Count > 0 ? traceElms[0].InnerText.Trim() : "",
                    };
                } else if(doc.DocumentElement.Name == "job") {
                    return JobData.FromXml(doc.DocumentElement);
                } 
                throw new JobRequestError
                {
                    Type = "JobUnknownResponse",
                    Description = "Unrecognized response",
                    Trace = "",
                };
            }
            catch (Exception e) {
                AddinLogger.Error(new Exception($"Error launching job", e));
                throw e;
            }
        }

        public async Task<JobData> CheckJobUpdate(JobData job)
        {
            try {
                var response = await connection.GetAsync(job.Href);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(await response.Content.ReadAsStringAsync());
                if (doc.DocumentElement == null) {
                    throw new Exception("Invalid response: expected root element 'job' not found.");
                }
                if (doc.DocumentElement.Name == "error") {
                    var descElms = doc.DocumentElement.GetElementsByTagName("description");
                    var traceElms = doc.DocumentElement.GetElementsByTagName("trace");
                    throw new JobRequestError
                    {
                        Type = "JobRequestError",
                        Description = descElms.Count > 0 ? descElms[0].InnerText.Trim() : "",
                        Trace = traceElms.Count > 0 ? traceElms[0].InnerText.Trim() : "",
                    };
                } else if (doc.DocumentElement.Name == "job") {
                    return JobData.FromXml(doc.DocumentElement);
                }
                throw new JobRequestError
                {
                    Type = "JobUnknownResponse",
                    Description = "Unrecognized response",
                    Trace = "",
                };
            }
            catch (Exception e) {
                AddinLogger.Error(new Exception($"Error in job", e));
                throw e;
            }
        }
        public async void DeleteJob(JobData job)
        {
            try {
                await connection.DeleteAsync(job.Href);
            }
            catch (Exception e) {
                AddinLogger.Error(new Exception($"Error while deleting job", e));
                throw e;
            }
        }

        public async Task<JobData> DownloadResults(JobData job, string targetFolder)
        {
            try {
                var namedResults = job.Results?.NamedResults ?? new List<NamedResult>();
                var downloadTasks = namedResults.Select(r =>
                    DownloadNamedResult(r, System.IO.Path.Combine(targetFolder, r.Nicename ?? r.Name))
                ).ToList();

                var downloadedNamedResults = await Task.WhenAll(downloadTasks);

                var newJob = new JobData() {
                    Type = job.Type,
                    JobId = job.JobId,
                    Priority = job.Priority,
                    Status = job.Status,
                    Log = job.Log,
                    Results = job.Results,
                    DownloadedFolder = job.DownloadedFolder,
                    Messages = job.Messages,
                    Progress = job.Progress,
                    Script = job.Script,
                    Nicename = job.Nicename,
                    ScriptHref = job.ScriptHref,
                    Href = job.Href,
                };
                newJob.DownloadedFolder = targetFolder;
                if (newJob.Results == null)
                    newJob.Results = new Results();
                newJob.Results.NamedResults = downloadedNamedResults.ToList();

                // Télécharger le log après les résultats
                return await DownloadJobLog(newJob, targetFolder);
            }
            catch (Exception e) {
                AddinLogger.Error(new Exception("Error downloading job results", e));
                return job;
            }

        }

        private async Task<NamedResult> DownloadNamedResult(NamedResult r, string targetUrl)
        {
            try {
                var response = await connection.GetAsync(r.Href);
                var buffer = await response.Content.ReadAsByteArrayAsync();
                List<string> files;
                if (r.MimeType == "application/zip")
                    files = await UnzipFile(buffer, targetUrl);
                else
                    files = await SaveFile(buffer, targetUrl);

                var filesUrls = files.Select(f => new Uri(f)).ToList();
                var newResult = new NamedResult()
                {
                    Desc = r.Desc,
                    Name = r.Name,
                    Nicename = r.Nicename,
                    MimeType = r.MimeType,
                    From = r.From,
                    Files = new List<ResultFile>(),
                };
                newResult.Href = targetUrl;
                newResult.Files = r.Files.Select(res =>
                {
                    var newResultFile = new ResultFile()
                    {
                        File = res.File,
                        Href = res.Href,
                        MimeType = res.MimeType,
                        Size = res.Size,
                    };
                    
                    var urlFound = filesUrls.FirstOrDefault(furl =>
                        res.File != null && furl.AbsoluteUri.EndsWith(res.File.Substring(targetUrl.Length))
                    );
                    if (urlFound != null) {
                        newResultFile.File = urlFound.AbsoluteUri;
                    }
                    return newResultFile;
                }).ToList();
                return newResult;
            }
            catch (Exception e) {
                // En cas d’erreur, retourner le résultat original
                AddinLogger.Error(new Exception("Error downloading named result", e));
                return r;
            }
        }

        private async Task<JobData> DownloadJobLog(JobData job, string targetFolder)
        {
            try {
                var jobTargetUrl = System.IO.Path.Combine(targetFolder, "job.log");
                var log = await connection.GetAsync(job.Log);
                await SaveFile(new System.Text.UTF8Encoding().GetBytes(await log.Content.ReadAsStringAsync()), jobTargetUrl);

                var newJob = new JobData()
                {
                    Type = job.Type,
                    JobId = job.JobId,
                    Priority = job.Priority,
                    Status = job.Status,
                    Log = job.Log,
                    Results = job.Results,
                    DownloadedFolder = job.DownloadedFolder,
                    Messages = job.Messages,
                    Progress = job.Progress,
                    Script = job.Script,
                    Nicename = job.Nicename,
                    ScriptHref = job.ScriptHref,
                    Href = job.Href,
                };

                newJob.Log = jobTargetUrl;
                if (newJob.Results == null)
                    newJob.Results = new Results();
                if (newJob.Results.NamedResults == null)
                    newJob.Results.NamedResults = new List<NamedResult>();
                return newJob;
            }
            catch (Exception e) {
                AddinLogger.Error(new Exception("Error downloading job log", e));
                return job;
            }
        }

        /// <summary>
        /// Portage from pipeline app unzipFile method : 
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="targetUrl"></param>
        /// <returns></returns>
        private static async Task<List<string>> UnzipFile(byte[] buffer, string targetDirectory)
        {
            var extractedFiles = new List<string>();
            try {
                if (!Directory.Exists(targetDirectory)) {
                    Directory.CreateDirectory(targetDirectory);
                }

                using (var memoryStream = new MemoryStream(buffer))
                using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Read)) {
                    foreach (var entry in archive.Entries) {
                        var destinationPath = System.IO.Path.Combine(targetDirectory, entry.FullName);

                        // Crée les sous-dossiers si besoin
                        var destinationDir = System.IO.Path.GetDirectoryName(destinationPath);
                        if (!string.IsNullOrEmpty(destinationDir) && !Directory.Exists(destinationDir)) {
                            Directory.CreateDirectory(destinationDir);
                        }

                        if (!string.IsNullOrEmpty(entry.Name)) // Ignore les dossiers
                        {
                            using (var entryStream = entry.Open())
                            using (var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write)) {
                                await entryStream.CopyToAsync(fileStream);
                            }
                            extractedFiles.Add(destinationPath);
                        }
                    }
                }
            }
            catch (Exception ex) {
                AddinLogger.Error(new Exception("Error unzipping file", ex));
                // returns an empty list in case of error (same as in the pipeline app)
                return new List<string>();
            }
            return extractedFiles;
        }

        /// <summary>
        /// Portage from pipeline app saveFile method : Save a downloaded buffer to a file on disk.
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="targetUrl"></param>
        /// <returns></returns>
        private static async Task<List<string>> SaveFile(byte[] buffer, string targetUrl)
        {
            var filePath = new Uri(targetUrl).LocalPath;
            var result = new List<string>();
            try {
                var directory = System.IO.Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
                    Directory.CreateDirectory(directory);
                }

                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)) {
                    await fileStream.WriteAsync(buffer, 0, buffer.Length);
                }
                result.Add(filePath);
            }
            catch (Exception e) {
                AddinLogger.Error(new Exception($"An error occurred writing file {filePath}", e));
                // returns an empty list in case of error (same as in the pipeline app)
            }
            return result;
        }



    }
}
