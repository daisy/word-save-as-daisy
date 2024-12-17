using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Daisy.SaveAsDAISY.Conversion.Events;

namespace Daisy.SaveAsDAISY.Conversion
{
    /// <summary>
    /// Abstract class for pipeline 2 scripts launch
    /// </summary>
    public abstract class Pipeline2Script : Script
    {
        public class JobException : Exception
        {
            public JobException(string message)
                : base(message) { }
        }

        protected Pipeline2Script(IConversionEventsHandler e)
            : base(e) { }

        public abstract string searchInputFromDirectory(DirectoryInfo inputDirectory);

        //IntPtr currentJob = IntPtr.Zero;

        /// <summary>
        /// executes script using jnet bridge
        /// </summary>
        /// <param name="inputPath">path to use for script main input</param>
        /// <param name="isQuite"></param>
        public override void ExecuteScript(string inputPath, bool isQuite)
        {
            try
            {
                TryExecuteOnAppWebservice(inputPath, isQuite);
            }
            catch (JobException e) {
                throw e;
            }
            catch (Exception e)
            {
                this.EventsHandler.onFeedbackMessageReceived(
                    this,
                    new DaisyEventArgs(
                        $"The script could not be launched on Pipeline App webservice:{e.Message}\r\n > Using addin embedded pipeline"
                    )
                );
                ExecuteOnEmbeddedPipeline(inputPath, isQuite);
                //throw new Exception("An error occured while executing the script " + this.Name, e);
            }
        }

        /// <summary>
        /// Run the pipeline 2 script on an embedded pipeline
        /// </summary>
        /// <param name="inputPath"></param>
        /// <param name="isQuite"></param>
        /// <exception cref="JobException"></exception>
        /// <exception cref="Exception"></exception>
        public void ExecuteOnEmbeddedPipeline(string inputPath, bool isQuite)
        {
            Pipeline2 pipeline = Pipeline2.Instance;
            Pipeline2.Instance.SetPipelineErrorListener(
                (message) =>
                {
                    this.EventsHandler.OnConversionError(
                        new Exception(
                            "Pipeline 2 returned the following error message :\r\n" + message
                        )
                    );
                }
            );
            Pipeline2.Instance.SetPipelineOutputListener(
                (message) =>
                {
                    this.EventsHandler.onFeedbackMessageReceived(this, new DaisyEventArgs(message));
                }
            );
            if (Parameters.ContainsKey("input") && (string)Parameters["input"].ParameterValue == "")
            {
                Parameters["input"].ParameterValue = inputPath;
            }
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            foreach (KeyValuePair<string, ScriptParameter> v in Parameters)
            {
                // avoid passing empty values
                if (
                    !v.Value.IsParameterRequired
                    && (
                        v.Value.ParameterDataType is StringDataType
                        || v.Value.ParameterDataType is PathDataType
                    )
                    && "" == (string)v.Value.ParameterValue
                )
                {
                    continue;
                }
                parameters[v.Value.Name] = v.Value.ParameterValue;
            }
            IntPtr currentJob = Pipeline2.Instance.Start(Name, parameters);

            if (currentJob != IntPtr.Zero)
            {
                bool checkStatus = true;
                List<string> errors;
                while (checkStatus)
                {
                    foreach (string message in pipeline.getNewMessages(currentJob))
                    {
                        this.EventsHandler.onFeedbackMessageReceived(
                            this,
                            new DaisyEventArgs("DP2 > " + message)
                        );
                    }
                    switch (pipeline.getStatus(currentJob))
                    {
                        case Pipeline2.JobStatus.IDLE:
                            break;
                        case Pipeline2.JobStatus.RUNNING:
                            break;
                        case Pipeline2.JobStatus.SUCCESS:
                            checkStatus = false;
                            break;
                        case Pipeline2.JobStatus.ERROR:
                            errors = pipeline.getErros(currentJob);
                            string errorMessage =
                                " DP2 > "
                                + this.NiceName
                                + " conversion job has finished in error :\r\n"
                                + string.Join("\r\n", errors);
                            throw new JobException(errorMessage);
                        case Pipeline2.JobStatus.FAIL:
                            // open jobs folder
                            errors = pipeline.getErros(currentJob);
                            string failedMessage =
                                " DP2 > "
                                + this.NiceName
                                + " conversion job failed :\r\n"
                                + string.Join("\r\n", errors);
                            throw new JobException(failedMessage);
                        default:
                            break;
                    }
                    System.Threading.Thread.Sleep(1000);
#if DEBUG
                    // Kill the instance and running jvm for pipeline debugging
                    try
                    {
                        //Pipeline2.KillInstance();
                    }
                    catch (Exception e)
                    {
                        throw;
                    }
                    //
#else
                    if (!isQuite && !string.IsNullOrEmpty(output))
                        System.Diagnostics.Process.Start(output);
#endif
                }
            }
            else
            {
                throw new Exception(
                    "DP2 > An unknown error occured while launching the script "
                        + this.Name
                        + " with the parameters "
                        + this.Parameters.Aggregate(
                            "",
                            (result, keyvalue) =>
                                result
                                + keyvalue.Value.Name
                                + "="
                                + keyvalue.Value.ParameterValue.ToString()
                                + "\r\n"
                        )
                );
            }
        }

        public void TryExecuteOnAppWebservice(string inputPath, bool isQuite)
        {
            Webservice engine = null;
            if (!PipelineApp.IsInstalled())
            {
                throw new Exception("DAISY pipeline app is not installed");
            }
            if (!PipelineApp.IsRunning())
            {
                // try launching daisy pipeline app in background
                // wait for it to start
                engine = PipelineApp.Start();
            }
            else
            {
                engine = PipelineApp.FindWebservice();
            }
            if (engine == null)
            {
                throw new Exception("Could not access the DAISY pipeline app webservice");
            }
            // Check if pipeline is alive
            Alive alive = engine.FetchAlive();
            if (!alive.alive)
            {
                throw new Exception("DAISY pipeline app webservice is not responding");
            }
            if (Parameters.ContainsKey("input") && (string)Parameters["input"].ParameterValue == "")
            {
                Parameters["input"].ParameterValue = inputPath;
            }
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            foreach (KeyValuePair<string, ScriptParameter> v in Parameters)
            {
                // avoid passing empty values
                if (
                    !v.Value.IsParameterRequired
                    && (
                        v.Value.ParameterDataType is StringDataType
                        || v.Value.ParameterDataType is PathDataType
                    )
                    && "" == (string)v.Value.ParameterValue
                )
                {
                    continue;
                }
                parameters[v.Value.Name] = v.Value.ParameterValue;
            }
            string outputFolder = Path.GetDirectoryName(
                Parameters["output"].ParameterValue.ToString()
            );
            JobRequest jobStart = JobRequest.fromScript(engine, this);
            JobData job = null;
            bool checkStatus = true;
            List<string> errors;
            while (checkStatus)
            {

                job = job == null ? engine.LaunchJob(jobStart) : engine.FetchJobDetails(job);
                string messages = job.GetNewMessages("App > ");
                if (messages.Length > 0)
                {
                    this.EventsHandler.onFeedbackMessageReceived(
                        this,
                        new DaisyEventArgs(messages)
                    );
                }
                switch (job.Status.ToLower())
                {
                    case "success":
                        checkStatus = false;
                        if (!Directory.Exists(outputFolder))
                        {
                            Directory.CreateDirectory(outputFolder);
                        }
                        // Download the result into the selected output folder
                        byte[] result = engine.FetchResults(job);
                        if (job.Results.MimeType == "application/zip")
                        {
                            UnzipByteArray(result, Parameters["output"].ParameterValue.ToString());
                        }
                        else
                        {
                            // TODO: should not be necessary (i'm expected a zip on this call)
                            // but need to check for mimetype and set the correct extension
                            File.WriteAllBytes(Path.Combine(outputFolder, "result"), result);
                        }
                        File.WriteAllText(
                            Path.Combine(outputFolder, "success.log"),
                            engine.FetchJobLog(job)
                        );
                        engine.DeleteJob(job);
                        break;
                    case "error":
                        string errorMessage =
                            " DP2 > "
                            + this.NiceName
                            + "conversion job has finished in error :\r\n"
                            + "Please check the log stored here:\r\n"
                            + Path.Combine(outputFolder, "error.log");
                        if (!Directory.Exists(outputFolder))
                        {
                            Directory.CreateDirectory(outputFolder);
                        }
                        File.WriteAllText(
                            Path.Combine(outputFolder, "error.log"),
                            engine.FetchJobLog(job)
                        );
                        engine.DeleteJob(job);
                        throw new JobException(errorMessage);
                    case "fail":
                        // open jobs folder
                        string failedMessage =
                            " DP2 > "
                            + this.NiceName
                            + " conversion job failed :\r\n"
                            + "Please check the log stored here:\r\n"
                            + Path.Combine(outputFolder, "error.log");
                        if (!Directory.Exists(outputFolder))
                        {
                            Directory.CreateDirectory(outputFolder);
                        }
                        File.WriteAllText(
                            Path.Combine(outputFolder, "error.log"),
                            engine.FetchJobLog(job)
                        );
                        engine.DeleteJob(job);
                        throw new JobException(failedMessage);
                    case "idle":
                    case "running":
                    default:
                        break;
                }
                System.Threading.Thread.Sleep(1000);

#if DEBUG
                // Kill the instance and running jvm for pipeline debugging
                try
                {
                    //Pipeline2.KillInstance();
                }
                catch (Exception e)
                {
                    throw;
                }
                //
#else
                if (!isQuite && !string.IsNullOrEmpty(output))
                    System.Diagnostics.Process.Start(output);
#endif
            }
        }

        public static void UnzipByteArray(byte[] zipBytes, string toBasePath)
        {
            if (!Directory.Exists(toBasePath))
            {
                Directory.CreateDirectory(toBasePath);
            }
            using (var memoryStream = new MemoryStream(zipBytes))
            {
                using (var zipArchive = new ZipArchive(memoryStream, ZipArchiveMode.Read))
                {
                    foreach (var entry in zipArchive.Entries)
                    {
                        using (var entryStream = entry.Open())
                        {
                            byte[] buffer = new byte[entry.Length];
                            entryStream.Read(buffer, 0, buffer.Length);
                            string relativePath = entry.FullName.StartsWith("result/")
                                ? entry.FullName.Substring("result/".Length)
                                : entry.FullName;
                            relativePath = relativePath.Replace("/", "\\");
                            string destination = Path.Combine(toBasePath, relativePath);
                            if (relativePath.EndsWith("/"))
                            {
                                Directory.CreateDirectory(destination);
                            }
                            else
                            {
                                if (!Directory.Exists(Path.GetDirectoryName(destination)))
                                {
                                    Directory.CreateDirectory(Path.GetDirectoryName(destination));
                                }
                                File.WriteAllBytes(destination, buffer);
                            }
                        }
                    }
                }
            }
        }
    }
}
