using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    // All members will be implicitly abstract and have no implementation
    public interface IUndoRedoCmd
    {
        IUndoRedoCmd Exec(Spreadsheet sheet);
    }

    // collection of functions to drive undoing and redoing
    public class UndoRedoCollection
    {
        public string CmdDescription;
        private IUndoRedoCmd[] m_commands;
        public UndoRedoCollection()
        {
        }
        public UndoRedoCollection(IUndoRedoCmd[] cmds, string description)
        {
            m_commands = cmds;
            CmdDescription = description;
        }
        public UndoRedoCollection(List<IUndoRedoCmd> cmds, string description)
        {
            m_commands = cmds.ToArray();
            CmdDescription = description;
        }

        // adds each command to cmdList with its associated description.
        // executes each of the listed command actions
        public UndoRedoCollection Exec(Spreadsheet ss)
        {
            List<IUndoRedoCmd> cmdList = new List<IUndoRedoCmd>();

            foreach (IUndoRedoCmd cmd in m_commands)
            {
                cmdList.Add(cmd.Exec(ss));
            }
            return new UndoRedoCollection(cmdList.ToArray(), this.CmdDescription);
        }
    }

}