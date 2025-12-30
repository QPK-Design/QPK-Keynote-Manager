namespace QPK_Keynote_Manager
{
    public enum FindReplaceScopeKind
    {
        Keynotes,
        SheetNames,
        ViewTitles
    }

    public class FindReplaceScope
    {
        public FindReplaceScopeKind Kind { get; }
        public string DisplayName { get; }

        public FindReplaceScope(FindReplaceScopeKind kind, string displayName)
        {
            Kind = kind;
            DisplayName = displayName;
        }
    }
}
