# fio Excel 自动化运行工具

面向 Linux，兼容 Python 3.11.6。脚本只使用 Python 标准库，直接读取 `.xlsx`，无需安装 openpyxl、pandas 或其他 Python 包。系统需要安装 **Flexible I/O Tester（fio）**，且该 fio 构建支持 `libaio`。

当前版本完成配置、生成和运行，不解析性能结果。最终性能输出是 fio 原生 plain text。

v1.4.0 新增主流程日志和当前用例进度记录；任一错误立即终止整批任务，不再支持出错后继续。

`examples/example_jobs.fio` 展示 4 个独立 job、每个 numjobs=2 的完整配置。它以假设的 1 TB 设备和 CPU 2–17 为例，仅供查看结构；实际运行由脚本读取设备容量后重新生成。`examples/example_two_jobs.fio` 使用假设的 100 GiB 设备，展示可以精确对齐时的 `size=42%`、`offset=42%` 写法。

## 快速使用

1. 克隆本仓库到 Linux 测试机。
2. 打开 `示例.xlsx`，修改 `Cases` 中的 `device`、`cpu_start`、`cpu_end` 和测试参数，保存为 `.xlsx`。
3. `device` 填 Linux 设备完整路径，例如 `/dev/nvme1n1` 或 `/dev/sdb`。示例中的盘符必须替换为你的实际测试设备。
4. 将 `--output` 放在测试设备之外的文件系统。

预览配置，不启动 fio 负载：

```bash
sudo python3.11 fio_excel_runner.py \
  --excel 示例.xlsx \
  --output ./fio_results \
  --dry-run
```

确认生成的 `jobs.fio` 和 `cpu_mapping.csv` 后正式运行：

```bash
sudo python3.11 fio_excel_runner.py \
  --excel 示例.xlsx \
  --output ./fio_results
```

示例默认只启用两个读用例，各运行 60 秒。若启用写入或混合读写用例，执行时加 `--allow-write`。写入会破坏目标设备前 84% 内的原有数据，包括可能位于开头的分区表；脚本不会自动卸载、格式化或清空磁盘。

```bash
sudo python3.11 fio_excel_runner.py \
  --excel 示例.xlsx \
  --output ./fio_results \
  --allow-write
```

若你的解释器命令是 `python3`，先用 `python3 --version` 确认版本，再替换示例命令。`sudo` 下也必须能找到这个解释器。

## Excel 填写方法

### Cases：每行一个用例

| 字段 | 含义 |
|---|---|
| case_id | 用例唯一标识，1–64 位英文字母、数字、下划线或短横线，以字母或数字开头 |
| enabled | `1` 运行、`0` 跳过；留空默认 `1` |
| device | 目标设备完整路径，如 `/dev/nvme1n1`、`/dev/sdb`；同一用例的所有 job 使用此设备 |
| rw | `read`、`write`、`randread`、`randwrite`、`rw`、`randrw` |
| bs | 一个固定块大小，如 `4k`、`128k`、`1m` 或字节数 `4096` |
| iodepth | 每个工作实例的队列深度，正整数 |
| jobs | 独立 job 配置段数量，生成 `[job_001]`、`[job_002]` 等 |
| numjobs | 每个独立 job 默认启动多少个实例，可单独覆盖 |
| cpu_start | Linux 逻辑 CPU 编号起点，包含本编号 |
| cpu_end | Linux 逻辑 CPU 编号终点，包含本编号，不小于起点 |
| runtime_s | `time_based=1` 时的运行秒数，正整数，留空默认 `60`；`time_based=0` 时必须留空 |
| rwmixread | 仅 `rw`/`randrw` 使用，读比例 0–100，留空默认 `50`；纯读/纯写留空 |
| notes | 备注，不传给 fio |
| time_based | `1` 按 `runtime_s` 定时运行；`0` 完成规划范围一次后退出；旧 Excel 没有此列时默认 `1` |
| timeout_s | watchdog 秒数；`time_based=0` 时必须填写；定时用例留空则使用 `runtime_s + --timeout-grace` |
| numa_mem_policy | 可选的 fio NUMA 内存策略，如 `bind:0`、`prefer:1`、`interleave:0-1`、`local`；留空不设置 |

`bs` 的 k/m/g 按 1024 换算；生成 job 文件时写为明确的字节数。第一版不支持 `bssplit`、块大小范围或在一个单元格填参数列表。要比较参数，请复制成多行用例。

盘符优先级：命令行 `--device` > Excel 当前行 `device`。不传 `--device` 时直接使用 Excel；Excel 留空且命令行未指定时会报错。填写不同设备的用例仍串行运行，并分别读取对应设备容量来计算前 84%。`run_info.json` 记录 Excel 原值、最终设备和来源，生成的 `jobs.fio` 使用最终设备作为 `filename`。

例如临时把所有选中用例统一改测 `/dev/sdb`：

```bash
sudo python3.11 fio_excel_runner.py --excel 示例.xlsx --device /dev/sdb --dry-run
```

旧版 Excel 只要已有 `device` 列即可继续使用；缺少 `time_based`、`timeout_s` 和 `numa_mem_policy` 时自动按原来的定时模式处理，并且不设置 NUMA 内存策略。

### JobOverrides：单独配置某个 job

用 `case_id` 和 `job_index` 指定要覆盖的 job。`job_index` 从 1 开始，不超过该用例的 `jobs`。可覆盖 `rw`、`bs`、`iodepth`、`numjobs`、`rwmixread`、`numa_mem_policy`，空白继承 Cases。

例如示例中的 `custom_jobs` 生成两个配置段：

- `job_001`：`read`、128 KiB、iodepth 16、numjobs 1。
- `job_002`：`randwrite`、4 KiB、iodepth 32、numjobs 3。

该用例共 4 个实例，实际数量为各 job 的 `numjobs` 之和。启用它时，在 Cases 中把 `enabled` 改为 `1`，并为执行添加 `--allow-write`。

`JobOverrides` 可以只保留表头，也可以删除整个工作表。`Guide` 是填写说明，脚本不会读取它。

生成 job 文件时，`rw`、`bs`、`iodepth`、`numjobs` 会分别判断所有独立 job 的最终有效值：某项完全一致时只写入 `[global]`，方便手工调整；某项存在 JobOverrides 差异时，该项写回每个 `[job_xxx]`。`offset`、`size`、`cpus_allowed` 始终保留在具体 job 中。

### 输入约束

- 保持 `Cases`、`JobOverrides` 工作表名称及第 1 行英文字段名不变，字段列顺序可以调整。
- 新用例直接往下添加；脚本按实际非空行读取，不依赖 Excel Table 的原有边界。
- 不要在配置表右侧添加其他列或文字；备注填入 `notes`，其他说明放到 `Guide`。
- 使用固定值，不支持公式。公式即使带 Excel 缓存结果也会被拒绝，避免读取旧值。
- 支持 Excel 的 shared strings、inline strings、数字单元格，以及空白的格式化行。
- 仅支持普通、未加密的 `.xlsx`，不支持 `.xls`、`.xlsm`、日期型输入或合并的配置单元格。
- 禁用行仍检查配置格式；CPU 是否可用只检查本次启用并选中的用例。
- 重复用例 ID、重复 job 覆盖、未知字段、缺失必要值和非法参数都会报错。

## 运行和绑核约定

不同用例按 Excel 行顺序串行运行，包括指向不同设备的用例。每个用例只启动一次 fio，文件中的所有独立 job 和它们的 `numjobs` 实例同时运行，没有 `stonewall`。

CPU 参数是 Linux **逻辑 CPU ID**，不是去除 SMT 后的物理 core 编号。可以在机器上用 `lscpu -e=CPU,CORE,SOCKET,NODE,ONLINE` 检查编号。

**每个独立 job 单独分配恰好 numjobs 个连续 CPU**，不再给所有 job 使用同一个全局 CPU 列表。例如该 job 的起始 CPU 是 251：

```ini
; numjobs=1 时
numjobs=1
cpus_allowed=251
cpus_allowed_policy=split

; numjobs=4 时
numjobs=4
cpus_allowed=251-254
cpus_allowed_policy=split
```

`cpus_allowed_policy=split` 保留，保证同一个 job 的每个实例各绑定区间内的一个 CPU，不在整个区间内自由迁移。

设输入闭区间有 C 个 CPU，各 job 的实例数为 n1、n2…，合计 W：

- 必须满足 `W <= C`；如果全部 job 需要的核数超过输入范围，生成计划时直接报错，不复用 CPU。
- 从 `cpu_start` 开始，按 job 顺序连续分配恰好 numjobs 个核，各 job 区间互不重叠。
- 分配足够的核后立即停止，输入范围内剩余的高编号 CPU 不参与本用例。

例如 CPU `1..100`、4 个 job、每个 numjobs=4，需要 16 个核，依次分配为 `1-4`、`5-8`、`9-12`、`13-16`，17-100 不使用。逐 job 配置的 numjobs 不同时，也按各自数量推进起始位置。分配结果保存在 `run_info.json` 的 `cpu_assignments`，并单独写入每个用例目录的 `cpu_binding.json`。

fio 的 split 仍按全局实例序号在该 job 的区间内取模。因此不同 numjobs 的 job 中，clone 0 不一定从区间第一个核开始，但同一个 job 的所有克隆恰好覆盖该 job 的全部分配核。

`cpu_mapping.csv` 是基于 fio split 规则生成的**预期分配**，不是运行时采样结果。当前版本不采集实际 CPU 调度轨迹，不修改 IRQ 亲和性，也不承诺各核 CPU 使用率或 IOPS 相同。按 CPU 编号连续分配并不等价于按 NUMA 或物理核拓扑分配。

整个输入闭区间必须在线且属于当前进程可用的 affinity/cpuset；不符合时明确报错，不会静默缩小范围。

### NUMA 内存绑定

CPU 亲和性和内存 NUMA 策略是两个独立设置：`cpus_allowed` 决定 fio worker 在哪些逻辑 CPU 上运行，`numa_mem_policy` 决定 fio I/O buffer 从哪些 NUMA node 分配。常用配置：

```text
numa_mem_policy=bind:0          # 严格绑定 node 0
numa_mem_policy=prefer:1        # 优先 node 1，内存不足时允许其他节点
numa_mem_policy=interleave:0-1  # 在 node 0、1 间交错分配
numa_mem_policy=local           # 使用 worker 执行时的本地节点
```

Cases 中的值是该用例所有独立 job 的默认策略；JobOverrides 可以逐 job 覆盖。一个 job 的所有 `numjobs` 实例共享这个策略。若不同实例需要不同内存节点，应拆成多个独立 job。

正式运行前会检查配置中的 NUMA node 是否在线。默认的 `--numa-backend auto` 会先检查 fio 是否原生认识 `numa_mem_policy`；如果当前 fio 构建没有 libnuma 支持，则自动从 job 文件移除该选项，并使用外层 `numactl` 启动 fio。例如 Excel 中的 `bind:2` 会执行为 `numactl --membind=2 fio ...`。可用 `numactl --hardware` 和 `cat /sys/devices/system/node/online` 查看节点。

可以显式选择后端：

```bash
# 推荐：原生 fio 可用就使用，否则自动回退到 numactl
sudo python3 fio_excel_runner.py --excel 示例.xlsx --output fio_results --numa-backend auto

# 即使 fio 原生支持，也强制使用 numactl
sudo python3 fio_excel_runner.py --excel 示例.xlsx --output fio_results --numa-backend numactl

# 只允许 fio 原生实现，不允许回退
sudo python3 fio_excel_runner.py --excel 示例.xlsx --output fio_results --numa-backend fio
```

外层 `numactl` 的内存策略会被整个 fio 进程及其所有 job 继承，因此同一个用例内所有 job 必须填写相同的 `numa_mem_policy`；不同用例可以使用不同 node。如果需要同一用例内每个 job 绑定不同 memnode，仍需使用带 `CONFIG_LIBNUMA` 的 fio 原生后端。CPU 绑定继续由 job 文件中的 `cpus_allowed` 控制。

可以先检查当前 fio 二进制是否包含 NUMA 支持：

```bash
fio --cmdhelp=numa_mem_policy
```

如果提示找不到该选项，即使 `fio --version` 是 3.34 或更高版本也不能原生使用，因为 NUMA 选项受编译时 `CONFIG_LIBNUMA` 控制。现在使用默认 `auto` 后端时会自动改用已安装的 numactl；只有显式指定 `--numa-backend fio` 才会因此终止。

## 磁盘范围与固定参数

**正式运行根据每个用例对应设备的实际容量计算，非固定容量。**前 84% 按独立 job 数量均分，各 job 访问的空间互不重叠；同一个 job 的 `numjobs` 克隆实例共享该 job 的区域。均分按独立 job 数量，不按 numjobs 加权。

两个 job 的目标范围是 `[0%, 42%)` 和 `[42%, 84%)`。结束地址不包含在该区域内，因此边界不会重复，也不需要从 43% 开始留下 1% 空隙。四个 job 则分别为 `[0%, 21%)`、`[21%, 42%)`、`[42%, 63%)`、`[63%, 84%)`。实际边界会按块对齐稍微向下调整。

脚本读取实际目标设备容量 D（字节），对于 N 个独立 job 计算：

```text
limit = floor(D × 84 / 100)
alignment = lcm(逻辑扇区大小, 本用例所有 job 的 bs)
units = floor(limit / alignment)
总区域大小 = units × alignment
第 i 个 job（i 从 0 开始）：
  offset = floor(units × i / N) × alignment
  end = floor(units × (i + 1) / N) × alignment
  size = end - offset
  访问区间 = [offset, end)，最后允许的字节地址是 end - 1
```

`offset` 和 `size` 写在各自的 `[job_xxx]` 下，不再放在 `[global]`。不设置 `offset_increment`，所以 `numjobs` 克隆不继续向后移动；增加 `jobs` 会重新均分区域。不同 job 的分区大小最多相差一个公共对齐单位。区域太小、无法给每个 job 分配至少一个对齐单位时直接报错。

**可以使用 `size=84%`。**fio 支持百分比大小，表示设备容量的相应比例，而不是结束地址。当前生成规则：

- 只有一个独立 job：输出 `offset=0`、`size=84%`；元数据记录 fio 向下块对齐后的实际字节边界。
- 多个独立 job：当 `84 / jobs` 为整数且所有边界都能精确对齐时，输出百分比，例如 `offset=42%`、`size=42%`。
- 其他情况：输出校验后的字节 offset/size，并附带 `Nominal disk range: [0%, 42%)` 和实际字节范围注释，保持可读性。

采用这个规则是因为 fio 百分比 offset 会向上对齐，而百分比 size 会向下对齐；直接拼接并不能对所有盘容量、块大小和 job 数量保证边界一致。比如 5 个 job 的理想长度是 16.8%，也不适合使用 fio 的整数百分比参数。脚本在这些情况下使用精确字节值，确保不重叠且不超过 84%。参见 [fio 的 size/offset 定义](https://fio.readthedocs.io/en/latest/fio_doc.html)。

每个 job 的计算结果保存在 `run_info.json` 的 `job_regions` 中，包括字节起点、长度、排他结束地址、名义百分比范围和实际 fio 参数。这个记录是生成配置的边界核算，不是运行时 I/O 跟踪。

总传输量可以因 `time_based` 循环运行而超过磁盘容量，这不表示访问地址越界。范围相对于本用例最终选定的设备计算；要测试整盘，请填整盘节点而不是某个分区。

以下参数固定，Excel 无法覆盖：

```ini
direct=1
ioengine=libaio
group_reporting=1
norandommap=1
randrepeat=0
ramp_time=0
```

其中 `norandommap=1` 等价于启用 `norandommap`。`time_based` 改为由每个 Cases 用例控制。

`time_based=1` 时生成：

```ini
time_based=1
runtime=60
```

`time_based=0` 时不生成 `time_based` 和 `runtime`，fio 完成规划范围一次后自然退出。比如示例中的 `precondition_write_84` 使用 `rw=write`、`jobs=1`、`numjobs=1`、`size=84%`，用于完整顺序写前 84%。该行默认禁用；启用后正式执行必须传 `--allow-write`。

另外固定 `allow_file_create=0`、`exitall_on_error=1`：不创建目标文件；一个 job 出错时停止本用例的其余 job。纯读用例额外使用 fio 的 `--readonly`。

## 输出归档

每次调用创建新的时间戳＋随机后缀目录，不覆盖旧结果：

```text
fio_results/
  20260922_210000_a1b2c3d4/
    runner.log                    # 主流程日志，与终端输出同步
    run_status.json               # 整批状态、当前阶段/用例序号、最终状态计数
    input.xlsx                    # 本次输入副本
    environment.json              # Python/fio 版本和实际设备信息
    001_seq_read_1m/
      jobs.fio                    # 完整 fio job 描述文件
      cpu_binding.json            # 本用例绑核策略、CPU 总数/清单和逐 job 区间
      cpu_mapping.csv             # 预期实例到 CPU 的对应关系
      cpu.csv                     # 每约 3 秒采集全部逻辑 CPU 的 /proc/stat 指标
      cpu_sampler.log             # CPU 采样器自身输出及异常
      memory.csv                  # fio 进程组、runner、系统和 NUMA 内存采样
      run_info.json               # 用例序号、配置、状态、命令、退出码、失败阶段/原因和时间
      parse_check.txt             # fio --parse-only 的输出
      result.txt                  # fio 原生 normal + JSON 组合输出，完整保留原始 stdout/stderr
      result.json                 # 从同一次 fio 运行提取的结构化结果，供统计工具读取
    002_randread_4k/
      ...
```

不启用每笔 I/O、BW、IOPS 或延迟时间序列日志。`group_reporting=1` 按你的要求保留，因此 plain text 中以 fio 分组汇总为准；不额外生成实例级性能报告。

`--dry-run` 只产生配置和元数据，不生成 `result.txt`/`parse_check.txt`，也不调用 fio。正式运行前，对所有选中用例做 fio `--parse-only` 检查，全部通过后才开始负载。

## 主流程日志和排查

`runner.log` 在读取 Excel 之前创建，每条日志带 UTC 时间、级别、阶段和用例位置，并同步输出到终端。例如：

```text
2026-09-22T14:00:00Z INFO [CONFIG_READ][suite] Reading Excel: /data/示例.xlsx
2026-09-22T14:00:01Z INFO [SELECT][suite] Selected 5 cases in worksheet order
2026-09-22T14:01:02Z INFO [CASE_RUN][case 2/5 id=randread_4k] Starting case; device=/dev/nvme1n1; workers=8; mode=time_based runtime=60s; watchdog=120s
2026-09-22T14:01:02Z INFO [CASE_RUN][case 2/5 id=randread_4k] Process started: pid=12345
2026-09-22T14:01:12Z INFO [CASE_RUN][case 2/5 id=randread_4k] Process running: pid=12345; elapsed=10.0s
```

这里的 `2/5` 表示本次选中的第 2 个用例、共 5 个，不包含禁用或未被 `--case` 选中的行。示例日志仅说明格式。

记录的阶段包括 `STARTUP`、`CONFIG_READ`、`SELECT`、`CPU_CHECK`、`NUMA_CHECK`、`DEVICE_CHECK`、`PLAN_GENERATE`、`WRITE_GUARD`、`FIO_CHECK`、`ARCHIVE`、`DEVICE_LOCK`、`PREFLIGHT`、`CASE_RECHECK`、`CASE_RUN`、`CASE_DONE`；预览完成时为 `PREVIEW_DONE`。设备、容量、CPU/NUMA/区域规划阶段、实际命令、子进程 PID、日志路径、退出码和中止信号都有记录。Python 异常保留 traceback。

fio 子进程等待期间每 10 秒输出一条存活等待日志。fio 使用 `--output-format=normal,json` 一次运行同时输出普通文本和 JSON；完整原始输出写入 `result.txt`，其中的 JSON 对象另存为 `result.json`，不会重复执行负载。

## 结果统计

统计工具会递归扫描结果根目录下的所有 `run_info.json`，以当前设计 Excel 的 `case_id` 和有效配置进行对照：

```bash
python3.11 fio_result_report.py \
  --excel 示例.xlsx \
  --results ./fio_results \
  --output 统计结果.xlsx
```

`--excel` 是可选参数。不提供 Excel 时不判断用例覆盖情况，只扫描已有归档或原始 fio 日志：

```bash
# 扫描整个结果目录（支持有或没有 run_info.json）
python3.11 fio_result_report.py --results ./fio_results --output 统计结果.xlsx

# 也可以直接提供单个 fio 原始日志
python3.11 fio_result_report.py --results result.txt --output 统计结果.xlsx
```

无 Excel 模式生成 `统计概览`、`运行明细`、`异常清单`。如果只有 `result.txt`/`result.json` 而没有 `run_info.json`，报告会提取可识别的性能汇总，将状态标记为“已收集（执行状态未知）”，并在“原始结果日志”列保留源文件路径；不会推测用例是否覆盖或执行是否成功。如果结果文件旁边同时存在 `cpu.csv` 和 `jobs.fio`，也会按 job 文件中的 CPU 绑定汇总 CPU 指标。

报表包含 `统计概览`、`用例覆盖`、`Job配置`、`运行明细` 和 `异常清单`。`用例覆盖` 每个用例一行，前面的配置列与输入 Excel 的 `Cases` 列一一对应，`rw`、`bs`、`iodepth`、`jobs`、`numjobs`、CPU 范围、运行模式等均为独立列，可直接使用 Excel 筛选；执行结果列后面是本用例绑定 CPU 的平均指标，再后面是 fio 性能汇总。多个 job 的某项有效参数完全相同时只显示一个值，不重复拼接；存在差异时显示“多值（见Job配置）”。`Job配置` 每个独立 job 一行，展开 `JobOverrides` 后的有效参数，便于筛选多 job 用例。`运行明细` 保留扫描到的全部历史记录，并给出原始结果和 CPU 采样日志路径。

CPU 汇总优先解析每个用例归档中的 `jobs.fio`，合并所有 job 的有效 `cpus_allowed` 并去重，再从 `cpu.csv` 中只选择对应的 `cpuN` 行。不会使用 Excel 的完整 `cpu_start..cpu_end` 直接推断，因为范围尾部可能有未使用的核。只有旧归档缺少 `jobs.fio` 时，才回退读取 `run_info.json` 的 `cpu_assignments`。

例如 `jobs.fio` 中四个 job 分别写着 `2-3`、`4-5`、`6-7`、`8-9`，报表只统计这 8 个逻辑 CPU，不会把输入范围中 10 及之后未绑定的 CPU 算进去。旧归档若有多个 job 重复绑定同一个 CPU，该 CPU 也只计一次。

CPU 汇总列包括：

- `cpu_binding_source`：绑定信息来源，正常为 `jobs.fio`，兼容回退时为 `run_info.cpu_assignments`。
- `cpu_bound_cpus`、`cpu_bound_cpu_count`：实际纳入统计的 CPU 列表和去重后的数量。
- `cpu_sample_count`：参与汇总的时间采样点数量。
- `cpu_sampled_cpus`：在 `cpu.csv` 中实际找到有效数据的绑定 CPU 数量。
- `cpu_avg_busy_pct`、`cpu_avg_user_pct`、`cpu_avg_system_pct`、`cpu_avg_iowait_pct`、`cpu_avg_idle_pct`：对绑定 CPU 和实际 `interval_s` 同时加权的平均百分比。
- `cpu_avg_core_equiv`：这些绑定 CPU 的平均忙碌核数。例如绑定 8 个逻辑 CPU 且 `cpu_avg_busy_pct=25%`，通常约等于 `cpu_avg_core_equiv=2`。

runner 仍采集系统可见的全部逻辑 CPU，因此同一份 `cpu.csv` 可以在统计时重新选择绑定核；报表中的上述平均值只代表本用例绑定集合。末尾原有的 `cpu_user_pct`、`cpu_sys_pct` 来自 fio 结果，反映 fio job 报告的进程 CPU 指标，二者不能直接混用。历史归档没有 `cpu.csv` 时 CPU 汇总列保持空白；CPU 日志、绑定信息缺失或内容损坏时，会写入“指标说明”并列入异常清单，不会静默改为全系统平均。

仓库提供 [结果统计样例](examples/结果统计样例.xlsx)，可以直接查看列顺序、筛选方式和两个示例用例的 CPU 汇总效果。

以下规则用于判断覆盖情况：

- `passed`、`failed`、`timeout`、`interrupted`、`error` 计入实际执行次数；dry-run 的 `planned` 不计入。
- 相同 `case_id` 的设备、CPU 范围、运行时间或任一 job 参数发生变化时，旧记录保留在运行明细，但当前配置标记为“配置已变更”，仍需重新执行。
- `enabled=0` 标记为“已禁用”；没有任何归档的启用用例标记为“从未执行”。
- 新版本归档从 `result.json` 读取性能数据。旧归档只有 `result.txt` 时会兼容解析常见 fio normal 输出；无法可靠识别的字段保持空白，并在异常清单中说明。

统计结果是用例级汇总。由于 job 文件固定使用 `group_reporting=1`，不输出实例级或独立 job 级性能表。

## 笛卡尔积用例生成

`fio_case_generator.py` 根据参数列表生成可直接交给 runner 的 Excel。下面的命令会生成 `3 × 3 × 4 = 36` 个用例：

```bash
python3.11 fio_case_generator.py \
  --rw read,write,randread \
  --bs 8K,128K,1M \
  --iodepth 1,2,4,8 \
  --device /dev/nvme1n1 \
  --cpu-start 2 --cpu-end 17 \
  --output 生成用例.xlsx
```

默认值为 `jobs=1`、`numjobs=1`、`runtime=60`。这三个参数也可以传入逗号分隔列表，加入笛卡尔积：

```bash
python3.11 fio_case_generator.py \
  --rw randread,randwrite --bs 4K,128K --iodepth 1,32 \
  --jobs 1,2 --numjobs 1,4 --runtime 30,60 \
  --device /dev/nvme1n1 --cpu-start 0 --cpu-end 15 \
  --case-prefix perf --output 生成用例.xlsx
```

生成器也可以为全部生成用例设置 NUMA 内存策略：

```bash
python3.11 fio_case_generator.py \
  --rw randread --bs 4K --iodepth 32 \
  --device /dev/nvme1n1 --cpu-start 0 --cpu-end 7 \
  --numa-mem-policy bind:0 --output NUMA用例.xlsx
```

生成器输出 `Cases` 和空的 `JobOverrides` 两个 sheet，并在写入后调用 runner 自己的 Excel 解析器做一次完整校验。默认最多生成 10000 个用例，可用 `--max-cases` 调整保护上限。对于 `rw`/`randrw`，可用 `--rwmixread` 指定读比例。

生成完整范围用例时关闭定时模式，并显式设置 watchdog：

```bash
python3.11 fio_case_generator.py \
  --rw write --bs 1M --iodepth 32 \
  --jobs 1 --numjobs 1 \
  --time-based 0 --timeout 7200 \
  --device /dev/nvme1n1 --cpu-start 2 --cpu-end 2 \
  --case-prefix precondition --output 预写用例.xlsx
```

可用以下命令跟踪进度，把目录替换成这次运行实际输出的归档目录：

```bash
tail -f fio_results/本次运行目录/runner.log
cat fio_results/本次运行目录/run_status.json
```

`run_status.json` 的 `stage`、`case_index`、`case_count`、`case_id` 表示当前阶段和用例。整批结束时还包含 `status`、`exit_code`、`finished_at`、各用例最终状态和 `counts`；失败时保留出错阶段、用例位置及 `error`。单个用例详情在各自的 `run_info.json` 中。

日志目录在配置校验前创建，所以 Excel 缺失、内容错误、设备检查失败也会留下日志。命令行参数解析错误，或输出目录/日志文件本身无法创建时，只能在终端报告。

## 其他命令

只运行指定的启用用例，可重复使用 `--case`：

```bash
sudo python3.11 fio_excel_runner.py --excel 示例.xlsx \
  --case randread_4k --output ./fio_results
```

**任一错误立即终止整批任务。**包括 fio 非零退出、超时、预检失败、设备/CPU 检查失败、启动子进程失败和 Python 异常。`--continue-on-error` 已移除，旧命令若仍携带此参数会直接报参数错误。

出错用例记录为 `failed`、`timeout`、`preflight_failed` 或 `error`，已完成用例保留原状态；尚未运行的用例标记为 `skipped` 并写入原因，不会自动重试或启动后续用例。如果是在全部用例的预检阶段失败，则所有负载都不启动，其他用例标记为 `skipped`。中断时当前用例记为 `interrupted`，其余未运行用例也标记为 `skipped`。

定时用例未填写 `timeout_s` 时，watchdog 上限为 `runtime_s + 60` 秒，包括 fio 启动开销，可以用 `--timeout-grace 120` 调整额外时间。填写了 `timeout_s` 时优先使用该值。`time_based=0` 必须填写 `timeout_s`，它只用于框架 watchdog，不会作为 fio 的 `runtime`。超时和 Ctrl+C 会先通知 fio 结束并等待报告落盘，再逐级终止；若必须强制杀死，`result.txt` 可能只含部分输出，状态会注明超时或中断。

没有连接测试设备时，可以离线预览（Windows/Linux 均可）：

```bash
python3.11 fio_excel_runner.py --excel 示例.xlsx \
  --dry-run \
  --device-size-bytes 1000000000000 --logical-sector-bytes 512 \
  --output ./preview_results
```

离线容量只允许配合 `--dry-run`，对所有选中设备使用同一个假设容量；正式运行一定分别读取实际设备容量。离线预览不验证设备存在性、CPU 在线情况或 fio 能力。

可以用 `--fio /usr/local/bin/fio` 指定 fio 可执行文件。

退出码：`0` 全部成功或预览成功；`1` fio 用例失败/超时；`2` 配置或运行环境错误；`130` 用户中断。

## 设备检查和使用边界

执行前检查目标是否为块设备，并拒绝已挂载的设备/子分区、活动 swap、带 device-mapper/RAID holder 的设备和 host-managed zoned 设备。系统盘保护会从 `/`、`/boot`、`/boot/efi`、`/usr`、`/var` 的挂载设备出发，沿分区父设备和 LVM/device-mapper/RAID 的 slave 关系追溯到底层物理盘。写用例命中系统盘时不会启动 fio，而是将该用例标记为 `skipped`，并在 `run_info.json` 中记录 `skip_code=system_disk` 和明确原因；`--allow-write` 不能绕过。每个用例开始前都会重新检查设备身份、容量、扇区大小、系统盘关系和 CPU 可用范围。

每个用例目录都会生成 `fio.cmd`，其中保存该用例准备执行的完整 workload 命令（包括 numactl、fio 路径和 `--readonly` 等参数），便于复核和排查。脚本仍以参数数组直接启动进程，不通过 shell 执行该文本文件。

正式 workload 运行期间还会每约 2 秒追加一次 `memory.csv`，包含 fio 进程组的进程数、RSS、VmHWM 汇总，runner RSS，系统 `MemAvailable/MemFree/SwapFree`，以及每个 NUMA node 的 `MemFree/MemUsed`。`run_info.json` 会记录 `estimated_io_buffer_bytes`，其估算公式为 `sum(numjobs × iodepth × bs)`；这是定位并发 buffer 风险的量级估算，不是 fio 总内存的严格上限。语法预检不会启动该监控。

每个正式 workload 还会并行启动项目内置的 `cpu_sample.py`，默认每 3 秒读取一次 `/proc/stat`。runner 不传 `--cpus`，因此会采集系统可见的全部逻辑 CPU，而不只采集 fio 绑定范围。每个用例目录保存 `cpu.csv` 和 `cpu_sampler.log`，采样间隔及采样器命令同时记录在 `run_info.json`。可以用 `--cpu-sample-interval 1` 调整间隔，或用 `--cpu-sampler /path/to/cpu_sample.py` 指定兼容脚本。采样仅覆盖正式 fio workload，不覆盖 `--parse-only` 预检。

同一设备的两个本工具实例通过 `/run/lock` 锁互斥。该锁不阻止其他软件访问磁盘；应在专用空闲测试盘上运行，也不要在运行期间重新挂载设备。挂载检查以当前 Linux mount namespace 为准，推荐直接在测试主机上运行。

## 已完成的验证

仓库的 `tests/test_runner.py` 包含自动测试，可运行：

```bash
python3.11 -B -m unittest discover -s tests -v
```

本次在 Windows / Python 3.12.14 环境完成 76 项自动测试，包括实际示例 `.xlsx` 读取、Excel 盘符和命令行覆盖优先级、公共参数自动提升、逐 job 覆盖、无 Excel 原始结果收集、全部逻辑 CPU 采样、`jobs.fio` 绑定核解析及加权汇总、缺少绑定信息时禁止回退全系统平均、内存/NUMA 采样、系统盘保护、从起始核连续且不复用的逐 job CPU 分配、独立绑核清单、84% 区域隔离、子进程超时及失败/中断状态等。分区检查覆盖 1、2、3、4、5、7、16、85 个 job 和不同容量，按 fio 的参数规则核算生成文件中的 offset/size。源码额外通过 Python 3.11 语法模式检查，仅使用 3.11 已有标准库 API。

结果统计样例的五张工作表已渲染检查，未发现公式错误或标题截断。尚未在 Python 3.11.6 本机或 Linux/libaio/真实块设备上执行集成测试；目标机正式运行会先进行 fio 参数解析检查。请先在目标机把一个只读用例改为短时长进行环境验证。

## fio 依据

- [fio 官方文档：job 文件、numjobs、范围和日志](https://fio.readthedocs.io/en/latest/fio_doc.html)
- [fio backend.c：split 使用全局 thread_number - 1](https://github.com/axboe/fio/blob/master/backend.c)
- [fio options.c：fio_cpus_split 按 CPU 数取模](https://github.com/axboe/fio/blob/master/options.c)
