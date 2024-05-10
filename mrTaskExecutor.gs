

function triggerПродолжитьВыполнениеЗадачь(trigerrInfo = "triggerПродолжитьВыполнениеЗадачь") {
  Logger.log(JSON.stringify(trigerrInfo));
  triggerLookCache();
  // return;
  let tasks = [].concat(
    new TaskListExportReestr().getTaskList(),
    new TaskListSyncLogist().getTaskList(),
  );

  let mrTaskExe = new MrTaskExecutor(getContext(), tasks, getContext().getScriptCache());
  try {
    mrTaskExe.triggerПродолжитьВыполнениеЗадачь();
  } catch (err) { mrErrToString(err); }

  triggerLookCache();
}



function triggerRemoveCache(trigerrInfo = "triggerRemoveCache") {
  Logger.log(JSON.stringify(trigerrInfo));
  getContext().getScriptCache();
  getContext().removeCache();
}

function triggerLookCache(trigerrInfo = "triggerLookCache") {
  Logger.log(JSON.stringify(trigerrInfo));
  Logger.log(JSON.stringify(getContext().getScriptCache()));

}


class MrTaskDataBaze {
  constructor() {
    this.class_key = undefined;
    this.name = undefined;
    this.off = false;
    this.param = {};
  }
}

class MrTaskExecutorBaze {
  constructor(sharedCache) {
    this.sharedCache = sharedCache;
    this.myCacheName = MrTaskExecutorBaze.name;
  }

  getMyCache() {
    if (!this.sharedCache[this.myCacheName]) {
      this.sharedCache[this.myCacheName] = new Object();
    }
    return this.sharedCache[this.myCacheName];
  }

  getCacheValue(key) {
    return this.getMyCache()[key];
  }


  setCacheValue(key, value) {
    // if (!this.sharedCache[this.myCacheName]) {
    //   this.sharedCache[this.myCacheName] = new Object();
    // }
    // return this.sharedCache[this.myCacheName][key];
    this.getMyCache()[key] = value;
    getContext().saveCache();
  }

  clearMyCache() {
    this.sharedCache[this.myCacheName] = undefined;
  }


  tryTask(self, colable_name, item) {
    try {
      return self[colable_name](item);
    } catch (err) {
      item.errors = [].concat(mrErrToString(err), item.errors);
    }
  }

}




class MrTaskExecutor extends MrTaskExecutorBaze {
  /** @param {MrTask} tasks */
  constructor(context, tasks, sharedCache, taskName) {
    super(sharedCache);
    if (!context) { throw new Error("context"); }
    /** @type {MrContext} */
    this.context = context;
    this.tasks = tasks;
    this.taskName = taskName;
    this.myCacheName = `${MrTaskExecutor.name}${taskName ? taskName : ""}`;
  }

  triggerПродолжитьВыполнениеЗадачь() {
    // this.context.getScriptCache().abc=2;
    let i = (this.getCacheValue("i") ? this.getCacheValue("i") : 0);
    if (i >= this.tasks.length) { i = 0; }
    // Logger.log(`i=${i} | this.tasks.length.=${this.tasks.length} `);
    // Logger.log(`${this.tasks} `);
    for (; i < this.tasks.length; i++) {
      this.taskExe(this.tasks[i]);
      // this.setCacheValue("i", i);
      // this.context.saveCache();
      if (!this.context.hasTime(1 / 24 / 60 * 28)) { Logger.log(`${this.myCacheName} triggerПродолжитьВыполнениеЗадачь | мало времени break`); break; }
    }


    if (i >= this.tasks.length) {
      i = undefined;
    }

    this.setCacheValue("i", i);
    this.context.saveCache();
  }

  /** @param {MrTask} task*/
  taskExe(task) {
    // Logger.log(`task=${JSON.stringify(task)}`);

    // let ss = MrTaskExe;
    let ss = this.context.getTaskExecutorsMap().get(task.exe_class);

    // Logger.log(`ss=${ss}   ${ss instanceof MrTaskExecutorBaze}`);
    // if (!(ss instanceof MrTaskExecutorBaze)) { return; }

    // Logger.log(ss.name)
    // Logger.log(new ss()[task.name]);


    new ss(this.context.getScriptCache())[task.name](task);
  }


}


