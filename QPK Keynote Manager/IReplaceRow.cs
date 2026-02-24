
using System.ComponentModel;

namespace QPK_Keynote_Manager
{
    public interface IReplaceRow
    {
        string Sheet { get; }

        event PropertyChangedEventHandler PropertyChanged;
    }
}
