using System.Diagnostics;

namespace VoicepeakWrapper;

public class Voicepeak
{
    private string exePath;

    public Voicepeak(string exePath)
    {
        this.exePath = exePath;

        if (!File.Exists(this.exePath))
        {
            throw new FileNotFoundException("VoicePeak not found. Please install VoicePeak and try again.");
        }
    }

    private async Task<string> RunCommandAsync(string cmd)
    {
        var process = new Process()
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = cmd,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            }
        };

        process.Start();
        string output = await process.StandardOutput.ReadToEndAsync();
        string error = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();

        if (!string.IsNullOrEmpty(error))
        {
            throw new Exception(error);
        }

        return output;
    }

    /// <summary>
    /// Constructs a command for the Voicepeak application.
    /// </summary>
    /// <param name="text">The text to be spoken. Mutually exclusive with textFile.</param>
    /// <param name="textFile">The path to a text file to be spoken. Mutually exclusive with text.</param>
    /// <param name="outputPath">The path where the output audio file will be saved.</param>
    /// <param name="narrator">The narrator to be used for speaking. Can be null.</param>
    /// <param name="emotions">The emotions to be applied during speech. Can be null.</param>
    /// <param name="speed">The speed of speech. Valid range is 50 to 200.</param>
    /// <param name="pitch">The pitch of speech. Valid range is -300 to 300.</param>
    /// <returns>A string representing the constructed command.</returns>
    /// <exception cref="ArgumentException">Thrown if both or neither text and textFile are provided.</exception>
    private string MakeSpeechCommand(string text = null, string textFile = null, string outputPath = null, Narrator narrator = null, Dictionary<string, int> emotions = null, int? speed = null, int? pitch = null)
    {
        List<string> commandParts = new List<string>();
    
        if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(textFile))
        {
            throw new ArgumentException("You can only specify one of text or textFile");
        }
        
        if (!string.IsNullOrEmpty(text))
        {
            commandParts.Add($"-s \"{text}\"");
        }
        else if (!string.IsNullOrEmpty(textFile))
        {
            commandParts.Add($"-t \"{textFile}\"");
        }
        else
        {
            throw new ArgumentException("You must specify either text or textFile");
        }
    
        if (!string.IsNullOrEmpty(outputPath))
        {
            commandParts.Add($"-o \"{outputPath}\"");
        }
    
        if (narrator != null)
        {
            commandParts.Add($"-n \"{narrator.Name}\"");
            if (emotions != null)
            {
                string emotionParams = string.Join(",", emotions.Select(kv => $"{kv.Key}={kv.Value}"));
                commandParts.Add($"-e {emotionParams}");
            }
        }
    
        if (speed.HasValue && speed >= 50 && speed <= 200)
        {
            commandParts.Add($"--speed {speed.Value}");
        }
    
        if (pitch.HasValue && pitch >= -300 && pitch <= 300)
        {
            commandParts.Add($"--pitch {pitch.Value}");
        }
    
        return string.Join(" ", commandParts);
    }
    
    /// <summary>
    /// Speaks the provided text using Voicepeak.
    /// </summary>
    /// <param name="text">The text to be spoken.</param>
    /// <param name="outputPath">The path where the output audio file will be saved.</param>
    /// <param name="narrator">The narrator to be used for speaking. Can be null.</param>
    /// <param name="emotions">The emotions to be applied during speech. Can be null.</param>
    /// <param name="speed">The speed of speech. Valid range is 50 to 200.</param>
    /// <param name="pitch">The pitch of speech. Valid range is -300 to 300.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    public async Task SayText(string text, string outputPath = null, Narrator narrator = null, Dictionary<string, int> emotions = null, int? speed = null, int? pitch = null)
    {
        string command = MakeSpeechCommand(text: text, outputPath: outputPath, narrator: narrator, emotions: emotions, speed: speed, pitch: pitch);
        await RunCommandAsync(command);
    }
    
    /// <summary>
    /// Speaks the text in the provided file using Voicepeak.
    /// </summary>
    /// <param name="textPath">The path to the text file to be spoken.</param>
    /// <param name="outputPath">The path where the output audio file will be saved.</param> 
    /// <param name="narrator">The narrator to be used for speaking. Can be null.</param>
    /// <param name="emotions">The emotions to be applied during speech. Can be null.</param>
    /// <param name="speed">The speed of speech. Valid range is 50 to 200.</param>
    /// <param name="pitch">The pitch of speech. Valid range is -300 to 300.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    public async Task SayTextFile(string textPath, string outputPath = "./output.wav", Narrator narrator = null, Dictionary<string, int> emotions = null, int? speed = null, int? pitch = null)
    {
        string command = MakeSpeechCommand(textFile: textPath, outputPath: outputPath, narrator: narrator, emotions: emotions, speed: speed, pitch: pitch);
        await RunCommandAsync(command);
    }
    
    /// <summary>
    /// Retrieves the list of available narrators.
    /// </summary>
    /// <returns>An array of Narrator objects.</returns>
    public async Task<Narrator[]> GetNarratorList()
    {
        var names = await GetNarratorNameList();
        var narrators = new List<Narrator>();
        foreach (var name in names)
        {
            var emotions = await GetEmotionList(name);
            narrators.Add(new Narrator(name, emotions));
        }
        return narrators.ToArray();
    }
    
    /// <summary>
    /// Retrieves the list of narrator names.
    /// </summary>
    /// <returns>An array of strings containing the names of available narrators.</returns>
    public async Task<string[]> GetNarratorNameList()
    {
        var output = await RunCommandAsync("--list-narrator");
        return output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
    }
    
    /// <summary>
    /// Retrieves the list of emotions available for a specific narrator.
    /// </summary>
    /// <param name="name">The name of the narrator.</param>
    /// <returns>An array of strings containing the names of available emotions for the given narrator.</returns>
    public async Task<string[]> GetEmotionList(string name)
    {
        var output = await RunCommandAsync($"--list-emotion \"{name}\"");
        return output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
    }
}