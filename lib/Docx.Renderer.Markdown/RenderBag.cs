namespace Docx.Renderer.Markdown;

public sealed class RenderBag : Dictionary<string, object>
{
    public T Get<T>(string key) => ContainsKey(key) ? (T)this[key] : default;

    public void AddOrReplace<T>(string key, T value)
    {
        Remove(key);
        Add(key,value);
    }
}