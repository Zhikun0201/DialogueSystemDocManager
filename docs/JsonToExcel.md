* 读取dlgjson文件，将其加载入ExcelData当中
* 特例：
   * Start没有名为Index的key
   * Selector没有名为Text的Key
   * Sequence没有名为Text的Key，有名为SpeechSequence的Key
   * End没有名为Text的Key
* 从StartNode开始，遍历每一个Node，制作Index与Children的字典，记作node_relations，格式为{Index: [children]}（Start的Index记作-1）
* 按照flow对node排序：
   * 获取node_relations
   * 新建一个字典，记作flow
   * 将开始节点的键值对放进去
   * 对于flow中最后一元素的值，检查值的各个元素是否出现在其他键的值当中
     * 出现的：跳到下一元素
     * 没出现的：将该元素从node_relations中pop，添加到flow当中
   * 直到node_relations为空，停止
* 计算缩进等级：
  * 检查一个节点是否有分支
  * 如果有分支，找到分支的目标节点，将目标节点的
