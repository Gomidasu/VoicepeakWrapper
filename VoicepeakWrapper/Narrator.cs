namespace VoicepeakWrapper;

public class Narrator
{
    public string Name { get; set; }
    public string[] Emotions { get; set; }

    public Narrator(string name, string[] emotions)
    {
        Name = name;
        Emotions = emotions;
    }
}