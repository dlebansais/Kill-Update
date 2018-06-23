using TaskbarIconHost;
using Microsoft.Win32.TaskScheduler;
using System;
using System.IO;

namespace SchedulerTools
{
    public static class Scheduler
    {
        #region Client Interface
        /// <summary>
        /// Adds a task that will launch a program every time someone logs in. The program must have privilege 'AsInvoker' and NOT 'Highest'.
        /// </summary>
        /// <param name="taskName">Task name, whatever you want that can be a file name</param>
        /// <param name="exeName">The full path to the program to launch</param>
        /// <param name="runLevel">The program privilege when launched</param>
        /// <returns>True if successful</returns>
        public static bool AddTask(string taskName, string exeName, TaskRunLevel runLevel, IPluginLogger logger)
        {
            try
            {
                // Remove forbidden characters since the name must not contain them.
                char[] InvalidChars = Path.GetInvalidFileNameChars();
                foreach (char InvalidChar in InvalidChars)
                    taskName = taskName.Replace(InvalidChar, ' ');

                // Create a task that launch a program when logging in.
                TaskService Scheduler = new TaskService();
                Trigger LogonTrigger = Trigger.CreateTrigger(TaskTriggerType.Logon);
                ExecAction RunAction = Microsoft.Win32.TaskScheduler.Action.CreateAction(TaskActionType.Execute) as ExecAction;
                RunAction.Path = exeName;

                // Try with a task name (mandatory on new versions of Windows)
                if (AddTaskToScheduler(Scheduler, taskName, LogonTrigger, RunAction, runLevel, logger))
                    return true;

                // Try without a task name (mandatory on old versions of Windows)
                if (AddTaskToScheduler(Scheduler, null, LogonTrigger, RunAction, runLevel, logger))
                    return true;
            }
            catch (Exception e)
            {
                logger.AddLog($"(from Scheduler.AddTask) {e.Message}");
            }

            return false;
        }

        /// <summary>
        /// Check if a particular task is active, by name.
        /// </summary>
        /// <param name="exeName"></param>
        /// <returns>True if active</returns>
        public static bool IsTaskActive(string exeName)
        {
            bool IsFound = false;
            EnumTasks(exeName, OnList, ref IsFound);
            return IsFound;
        }

        /// <summary>
        /// Remove an active task.
        /// </summary>
        /// <param name="exeName">Task name</param>
        /// <param name="isFound">True if found (and removed), false if not found</param>
        public static void RemoveTask(string exeName, out bool isFound)
        {
            isFound = false;
            EnumTasks(exeName, OnRemove, ref isFound);
        }
        #endregion

        #region Implementation
        private delegate void EnumTaskHandler(Task task, ref bool returnValue);

        private static void OnList(Task task, ref bool returnValue)
        {
            Trigger LogonTrigger = task.Definition.Triggers[0];
            if (LogonTrigger.Enabled)
                returnValue = true;
        }

        private static void OnRemove(Task task, ref bool returnValue)
        {
            TaskService Scheduler = task.TaskService;
            TaskFolder RootFolder = Scheduler.RootFolder;
            RootFolder.DeleteTask(task.Name, false);
        }

        private static void EnumTasks(string exeName, EnumTaskHandler handler, ref bool returnValue)
        {
            string ProgramName = Path.GetFileName(exeName);

            try
            {
                TaskService Scheduler = new TaskService();

                foreach (Task t in Scheduler.AllTasks)
                {
                    try
                    {
                        TaskDefinition Definition = t.Definition;
                        if (Definition.Actions.Count != 1 || Definition.Triggers.Count != 1)
                            continue;

                        ExecAction AsExecAction;
                        if ((AsExecAction = Definition.Actions[0] as ExecAction) == null)
                            continue;

                        if (!AsExecAction.Path.EndsWith(ProgramName) || Path.GetFileName(AsExecAction.Path) != ProgramName)
                            continue;

                        handler(t, ref returnValue);
                        return;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        private static bool AddTaskToScheduler(TaskService scheduler, string taskName, Trigger logonTrigger, ExecAction runAction, TaskRunLevel runLevel, IPluginLogger logger)
        {
            try
            {
                Task task = scheduler.AddTask(taskName, logonTrigger, runAction);
                task.Definition.Principal.RunLevel = runLevel;

                Task newTask = scheduler.RootFolder.RegisterTaskDefinition(taskName, task.Definition, TaskCreation.CreateOrUpdate, null, null, TaskLogonType.None, null);
                return true;
            }
            catch (Exception e)
            {
                logger.AddLog($"(from Scheduler.AddTaskToScheduler) {e.Message}");
            }

            return false;
        }
        #endregion
    }
}
