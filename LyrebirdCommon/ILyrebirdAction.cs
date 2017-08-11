using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lyrebird
{
    public abstract class LyrebirdAction
    {
        public abstract Guid CommandGuid { get; }
    }
}
