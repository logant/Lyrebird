using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lyrebird
{
    public interface ILyrebirdAction
    {
        Guid CommandGuid { get; }
    }
}
