param(
  [string]$SummaryPath = '.\artifacts\cis_analysis\cis_group_summary.csv',
  [string]$GroupedPath = '.\artifacts\cis_analysis\cis_group_user_agg.csv',
  [string]$OverlapPath = '.\artifacts\cis_analysis\cis_group_overlap.csv',
  [string]$OutputPath = '.\artifacts\cis_analysis\cis_visual_report.html'
)

$ErrorActionPreference = 'Stop'

$summaryRows = Import-Csv -LiteralPath $SummaryPath
$groupedRows = Import-Csv -LiteralPath $GroupedPath
$overlapRows = Import-Csv -LiteralPath $OverlapPath

$summaryMap = @{}
foreach ($row in $summaryRows) { $summaryMap[$row.segment_group] = $row }
$groupMap = @{}
foreach ($grp in ($groupedRows | Group-Object segment_group)) { $groupMap[$grp.Name] = $grp.Group }
$overlapMap = @{}
foreach ($row in $overlapRows) { $overlapMap["$($row.left_group)|$($row.right_group)"] = [double]$row.overlap_user_count }

$order = @('静默大户','稳定存量','价值大户','高频高风险','新用户')
$totalUsers = $groupedRows.Count

function Pct-Value($num, $den) {
  if ($null -eq $den -or [double]$den -eq 0) { return 0 }
  return [int][math]::Round((100.0 * [double]$num) / [double]$den, 0)
}
function Fmt-Pct($num, $den) {
  return ('{0}%' -f (Pct-Value $num $den))
}
function Fmt-K($value) {
  if ([string]::IsNullOrWhiteSpace([string]$value)) { return '0k' }
  $n = [double]$value
  if ($n -eq 0) { return '0k' }
  if ([math]::Abs($n) -lt 1000) { return '&lt;1k' }
  return ('{0}k' -f ([int][math]::Round($n / 1000.0, 0)))
}
function Fmt-Asset($value) {
  if ([string]::IsNullOrWhiteSpace([string]$value)) { return '0' }
  $n = [double]$value
  if ($n -gt 0 -and $n -lt 1) { return '&lt;1' }
  return ([int][math]::Round($n, 0)).ToString()
}
function Fmt-Int($value) {
  if ([string]::IsNullOrWhiteSpace([string]$value)) { return '0' }
  return ([int][math]::Round([double]$value, 0)).ToString()
}
function Lev-Bucket([string]$v) {
  if ([string]::IsNullOrWhiteSpace($v)) { return '未知' }
  $d = [double]$v
  if ($d -le 1) { return '1x' }
  if ($d -le 5) { return '2-5x' }
  if ($d -le 10) { return '6-10x' }
  return '10x+'
}
function Get-TopHtml($rows, [string]$field, [int]$limit, [double]$den, [switch]$ExcludeDash) {
  $filtered = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.$field) })
  if ($ExcludeDash) { $filtered = @($filtered | Where-Object { $_.$field -ne '-' }) }
  $top = @($filtered | Group-Object -Property $field | Sort-Object -Property @{Expression='Count';Descending=$true}, @{Expression='Name';Descending=$false} | Select-Object -First $limit)
  if ($top.Count -eq 0) { return "<span class='muted'>暂无有效信号</span>" }
  $chips = foreach ($item in $top) {
    "<span class='chip'><strong>$($item.Name)</strong><em>$(Fmt-Pct $item.Count $den)</em></span>"
  }
  return ($chips -join '')
}
function Bar-Section($items, $title, $valueField, $displayField) {
  $max = ($items | Measure-Object -Property $valueField -Maximum).Maximum
  if (-not $max) { $max = 1 }
  $rows = foreach ($item in $items) {
    $value = [double]$item.$valueField
    if ($value -eq 0) { $width = 0 } else { $width = [math]::Max(6, [int][math]::Round(100 * $value / $max, 0)) }
    "<div class='bar-row'><div class='bar-label'>$($item.segment_group)</div><div class='bar-track'><div class='bar-fill' style='width:${width}%'></div></div><div class='bar-value'>$($item.$displayField)</div></div>"
  }
  return "<section class='panel bars'><h3>$title</h3>$($rows -join '')</section>"
}

$largestGroup = $order | Sort-Object { -[int]$groupMap[$_].Count } | Select-Object -First 1
$highestTradeGroup = $order | Sort-Object { -[double]$summaryMap[$_].total_trade_amount_median } | Select-Object -First 1
$riskRows = @($groupMap['高频高风险'])
$risk10Share = Pct-Value (@($riskRows | Where-Object { (Lev-Bucket $_.leverage_primary) -eq '10x+' }).Count) $riskRows.Count

$thesisBlocks = @(
  "<article class='thesis-card'><div class='thesis-value'>$(Fmt-Pct $groupMap[$largestGroup].Count $totalUsers)</div><div class='thesis-title'>$largestGroup 占比</div><p>当前 CIS 结构首先是存量盘，运营第一优先级不是扩量，而是把沉默存量分层唤回。</p></article>",
  "<article class='thesis-card'><div class='thesis-value'>$(Fmt-K $summaryMap[$highestTradeGroup].total_trade_amount_median)</div><div class='thesis-title'>$highestTradeGroup 交易额中位数</div><p>高价值交易集中在小盘群体，适合用重点维护、费率和服务密度换取稳定产出。</p></article>",
  "<article class='thesis-card'><div class='thesis-value'>${risk10Share}%</div><div class='thesis-title'>高频高风险 10x+ 杠杆占比</div><p>这一组规模不大，但风险特征非常集中，激励和风控必须拆开管理。</p></article>"
)

$groupCards = foreach ($name in $order) {
  $rows = @($groupMap[$name])
  $summary = $summaryMap[$name]
  $share = Pct-Value $rows.Count $totalUsers
  $countriesHtml = Get-TopHtml $rows 'geo_country_primary' 3 $rows.Count
  $pairRows = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.pair_30d_primary) -and $_.pair_30d_primary -ne '-' })
  $pairCoverage = Pct-Value $pairRows.Count $rows.Count
  if ($name -eq '新用户') {
    $pairBlock = "<div class='signal-block'><label>币对信号</label><div class='detail-note'>新用户不单列币对偏好，避免把噪声当成可执行信号。</div></div>"
  } else {
    $pairsHtml = Get-TopHtml $pairRows 'pair_30d_primary' 3 $pairRows.Count
    $pairBlock = "<div class='signal-block'><label>Top 3 币对</label><div class='chips'>$pairsHtml</div><div class='detail-note'>有效币对覆盖 ${pairCoverage}% · 已剔除空值和 '-'。</div></div>"
  }
  $stance = switch ($name) {
    '静默大户' { '规模最大，但交易额中位数低；优先做资产残留分层和定向唤回，而不是全量刺激。' }
    '稳定存量' { '这是第二大的经营主池，活跃天数最高，适合承接会员权益、理财和复购任务。' }
    '价值大户' { '占比不高，但交易额中位数最高，重点应放在高价值维护而不是规模扩张。' }
    '高频高风险' { '小盘高杠杆，先设风险边界，再决定是否给交易激励。' }
    default { '先完成首周激活和首单引导，不放大噪声很大的偏好信号。' }
  }
  $accent = switch ($name) {
    '静默大户' { 'tone-silent' }
    '稳定存量' { 'tone-stable' }
    '价值大户' { 'tone-value' }
    '高频高风险' { 'tone-risk' }
    default { 'tone-new' }
  }
  "<article class='segment-card $accent'><div class='segment-top'><div><h3>$name</h3><p>$stance</p></div><div class='share-pill'>${share}%</div></div><div class='metric-grid'><div><span>群体占比</span><strong>${share}%</strong></div><div><span>交易额中位数</span><strong>$(Fmt-K $summary.total_trade_amount_median)</strong></div><div><span>资产中位数</span><strong>$(Fmt-Asset $summary.latest_asset_median)</strong></div><div><span>活跃天数中位数</span><strong>$(Fmt-Int $summary.active_days_median)d</strong></div></div><div class='signal-block'><label>Top 3 国家</label><div class='chips'>$countriesHtml</div></div>$pairBlock</article>"
}

$shareItems = foreach ($name in $order) { [pscustomobject]@{ segment_group = $name; value = (Pct-Value $groupMap[$name].Count $totalUsers); label = ((Pct-Value $groupMap[$name].Count $totalUsers).ToString() + '%') } }
$tradeItems = foreach ($name in $order) { [pscustomobject]@{ segment_group = $name; value = [double]$summaryMap[$name].total_trade_amount_median; label = (Fmt-K $summaryMap[$name].total_trade_amount_median) } }
$assetItems = foreach ($name in $order) { [pscustomobject]@{ segment_group = $name; value = [double]$summaryMap[$name].latest_asset_median; label = (Fmt-Asset $summaryMap[$name].latest_asset_median) } }
$activeItems = foreach ($name in $order) { [pscustomobject]@{ segment_group = $name; value = [double]$summaryMap[$name].active_days_median; label = ((Fmt-Int $summaryMap[$name].active_days_median) + 'd') } }
$corePanels = @(
  (Bar-Section $shareItems '群体结构占比' 'value' 'label'),
  (Bar-Section $tradeItems '交易额中位数' 'value' 'label'),
  (Bar-Section $assetItems '资产中位数' 'value' 'label'),
  (Bar-Section $activeItems '活跃天数中位数' 'value' 'label')
)

$countryFocus = foreach ($name in $order) {
  $rows = @($groupMap[$name])
  $countriesHtml = Get-TopHtml $rows 'geo_country_primary' 3 $rows.Count
  "<article class='focus-card'><h4>$name</h4><div class='chips'>$countriesHtml</div><div class='detail-note'>按群体内部占比展示。</div></article>"
}
$pairFocus = foreach ($name in $order | Where-Object { $_ -ne '新用户' }) {
  $rows = @($groupMap[$name] | Where-Object { -not [string]::IsNullOrWhiteSpace($_.pair_30d_primary) -and $_.pair_30d_primary -ne '-' })
  $coverage = Pct-Value $rows.Count $groupMap[$name].Count
  $pairsHtml = Get-TopHtml $rows 'pair_30d_primary' 3 $rows.Count
  "<article class='focus-card'><h4>$name</h4><div class='chips'>$pairsHtml</div><div class='detail-note'>有效币对覆盖 ${coverage}% · 已剔除空值和 '-'。</div></article>"
}

$overlapHeader = $order | ForEach-Object { "<th>$_</th>" }
$overlapRowsHtml = foreach ($left in $order) {
  $cells = foreach ($right in $order) {
    $pct = Pct-Value $overlapMap["$left|$right"] $groupMap[$left].Count
    $alpha = [math]::Max(0.08, [math]::Min(0.92, $pct / 100.0))
    "<td style='background:rgba(22,96,102,$alpha)'>$pct%</td>"
  }
  "<tr><th>$left</th>$($cells -join '')</tr>"
}

$strategyBlocks = @(
  "<article class='action-card'><h3>静默大户</h3><p>先按资产残留和国家做分层唤回，避免对低交易中位数的大盘人群做统一补贴。</p></article>",
  "<article class='action-card'><h3>稳定存量</h3><p>把预算投向权益、理财、会员和复购机制，这组既有规模，也有最高的活跃天数中位数。</p></article>",
  "<article class='action-card'><h3>价值大户</h3><p>优先做高价值维护和本地化服务，目标是稳住交易深度，而不是追求更大覆盖。</p></article>",
  "<article class='action-card'><h3>高频高风险</h3><p>单独管理，先设置杠杆和异常交易边界，再决定是否给予活动刺激。</p></article>",
  "<article class='action-card'><h3>新用户</h3><p>只做首周激活和首单转化，不把噪声很高的币对信号当成投放依据。</p></article>"
)

$html = @"
<!doctype html>
<html lang='zh-CN'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>CIS 用户分群决策看板</title>
<style>
:root{--bg:#f5f1e8;--paper:#fffdf8;--ink:#162226;--muted:#67767c;--line:#ddd2c3;--teal:#166066;--sand:#d6b98b;--amber:#c8782a;--shadow:0 18px 48px rgba(22,34,38,.08)}
*{box-sizing:border-box}body{margin:0;font-family:'Segoe UI','PingFang SC','Microsoft YaHei',sans-serif;color:var(--ink);background:radial-gradient(circle at top right,#efe6d5 0,#f5f1e8 38%,#efeae1 100%)}
body:before{content:'';position:fixed;inset:0;pointer-events:none;background-image:linear-gradient(rgba(255,255,255,.28) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.28) 1px,transparent 1px);background-size:28px 28px;opacity:.25}
main{position:relative;max-width:1380px;margin:0 auto;padding:32px 20px 64px}.hero{display:grid;grid-template-columns:1.2fr .8fr;gap:20px;align-items:stretch}.hero-panel,.hero-side{background:var(--paper);border:1px solid rgba(22,34,38,.06);border-radius:28px;box-shadow:var(--shadow)}
.hero-panel{padding:30px;background:linear-gradient(135deg,rgba(22,96,102,.95),rgba(30,54,57,.96));color:#f8f5ee;overflow:hidden;position:relative}.hero-panel:after{content:'';position:absolute;right:-60px;top:-40px;width:220px;height:220px;border-radius:50%;background:radial-gradient(circle,rgba(214,185,139,.45),rgba(214,185,139,0) 70%)}
.eyebrow{font-size:12px;letter-spacing:.18em;text-transform:uppercase;opacity:.76}.hero-panel h1{margin:12px 0 10px;font-size:42px;line-height:1.08}.hero-panel p{max-width:820px;font-size:17px;line-height:1.8;color:rgba(248,245,238,.86)}.hero-note{margin-top:18px;padding-top:16px;border-top:1px solid rgba(255,255,255,.18);font-size:14px;color:rgba(248,245,238,.78)}
.hero-side{padding:22px;display:grid;gap:14px;background:linear-gradient(180deg,#fffdf8,#f7f0e5)}.thesis-card{padding:16px 18px;border-radius:20px;background:#fff;box-shadow:0 12px 30px rgba(22,34,38,.06)}.thesis-value{font-size:34px;font-weight:700;color:var(--teal)}.thesis-title{margin-top:4px;font-size:15px;font-weight:700}.thesis-card p{margin:8px 0 0;color:var(--muted);line-height:1.65;font-size:14px}
.section{margin-top:26px}.section-head{display:flex;justify-content:space-between;align-items:end;gap:16px;margin-bottom:14px}.section-head h2{margin:0;font-size:26px}.section-head p{margin:0;color:var(--muted);max-width:760px;line-height:1.7}
.segment-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:18px}.segment-card{background:var(--paper);border-radius:24px;padding:20px;border:1px solid rgba(22,34,38,.06);box-shadow:var(--shadow)}.segment-top{display:flex;justify-content:space-between;gap:16px;align-items:start}.segment-top h3{margin:0;font-size:22px}.segment-top p{margin:8px 0 0;color:var(--muted);line-height:1.7;font-size:14px}.share-pill{padding:8px 12px;border-radius:999px;font-weight:700;font-size:18px;background:#ecf5f5;color:var(--teal);min-width:72px;text-align:center}
.metric-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin-top:18px}.metric-grid div{padding:12px 14px;border-radius:18px;background:#fff}.metric-grid span{display:block;font-size:12px;color:var(--muted);margin-bottom:6px}.metric-grid strong{font-size:22px}
.signal-block{margin-top:18px}.signal-block label{display:block;margin-bottom:8px;font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}.chips{display:flex;flex-wrap:wrap;gap:8px}.chip{display:inline-flex;align-items:center;gap:8px;padding:8px 10px;border-radius:999px;background:#f4eee5;border:1px solid rgba(22,34,38,.06);font-size:13px}.chip strong{font-weight:700}.chip em{font-style:normal;color:var(--teal);font-weight:700}.detail-note{margin-top:8px;font-size:12px;color:var(--muted)}.muted{font-size:13px;color:var(--muted)}
.tone-silent{border-top:5px solid #aa6c3d}.tone-stable{border-top:5px solid #1f7d72}.tone-value{border-top:5px solid #154b82}.tone-risk{border-top:5px solid #9c3d2b}.tone-new{border-top:5px solid #8a7d63}
.panel-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:18px}.panel{background:var(--paper);border-radius:24px;padding:20px;border:1px solid rgba(22,34,38,.06);box-shadow:var(--shadow)}.panel h3{margin:0 0 14px;font-size:20px}.bar-row{display:grid;grid-template-columns:110px 1fr 68px;gap:10px;align-items:center;margin:12px 0}.bar-label{font-size:13px}.bar-track{height:14px;background:#eadfce;border-radius:999px;overflow:hidden}.bar-fill{height:100%;background:linear-gradient(90deg,var(--amber),var(--teal));border-radius:999px}.bar-value{text-align:right;font-weight:700;font-size:13px}
.focus-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:16px}.focus-card{background:var(--paper);border-radius:20px;padding:18px;border:1px solid rgba(22,34,38,.06);box-shadow:var(--shadow)}.focus-card h4{margin:0 0 12px;font-size:18px}
.matrix-wrap{overflow:auto}.matrix{width:100%;border-collapse:separate;border-spacing:6px}.matrix th,.matrix td{padding:12px 10px;border-radius:14px;text-align:center;font-size:13px}.matrix th{background:#efe7d9}.matrix td{color:#fff;font-weight:700}
.actions{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:16px}.action-card{background:linear-gradient(180deg,#fffdf9,#f8f0e4);border-radius:22px;padding:20px;border:1px solid rgba(22,34,38,.06);box-shadow:var(--shadow)}.action-card h3{margin:0 0 10px;font-size:20px}.action-card p{margin:0;color:var(--muted);line-height:1.75}
@media (max-width:980px){.hero{grid-template-columns:1fr}.hero-panel h1{font-size:34px}}@media (max-width:640px){main{padding:22px 14px 44px}.segment-top{flex-direction:column}.share-pill{min-width:0}.metric-grid{grid-template-columns:1fr}.bar-row{grid-template-columns:84px 1fr 58px}}
</style>
</head>
<body>
<main>
  <section class='hero'>
    <article class='hero-panel'>
      <div class='eyebrow'>CIS Segmentation Decision View</div>
      <h1>当前重点不是扩盘，而是把存量分层经营清楚。</h1>
      <p>从结构上看，静默大户已经决定了页面的主问题；从价值上看，真正需要重点维护的是价值大户；从风险上看，高频高风险必须单独管理。下面的页面不再罗列绝对人数，只保留占比、关键中位数和可直接支持动作判断的偏好信号。</p>
      <div class='hero-note'>数据口径：工作簿最后修改时间为 2026-03-09，本页仅展示占比、交易额中位数（k）和资产中位数（整数），避免绝对人数干扰判断。</div>
    </article>
    <aside class='hero-side'>$(($thesisBlocks -join ''))</aside>
  </section>
  <section class='section'>
    <div class='section-head'><h2>结构与经营角色</h2><p>先看结构，再决定资源。这里每张卡只回答四个问题：它占多大盘子、交易中位数有多高、资产沉淀多不多、应该优先看哪些国家和币对。</p></div>
    <div class='segment-grid'>$(($groupCards -join ''))</div>
  </section>
  <section class='section'>
    <div class='section-head'><h2>核心差异</h2><p>这四组指标分别回答结构、价值、资产和活跃度，避免把页面做成“所有数字都堆在一起”的海报。</p></div>
    <div class='panel-grid'>$(($corePanels -join ''))</div>
  </section>
  <section class='section'>
    <div class='section-head'><h2>偏好信号</h2><p>国家信号按群体内部占比展示；币对信号只保留有效样本，并对静默大户剔除 '-'，对新用户不再单列币对偏好。</p></div>
    <div class='panel-grid'><section class='panel'><h3>国家焦点</h3><div class='focus-grid'>$(($countryFocus -join ''))</div></section><section class='panel'><h3>有效币对偏好</h3><div class='focus-grid'>$(($pairFocus -join ''))</div></section></div>
  </section>
  <section class='section'>
    <div class='section-head'><h2>群体重叠</h2><p>矩阵中的百分比表示：某个行群体中，有多少比例也同时落在列群体。这样比绝对人数更适合判断运营边界是否需要拆开。</p></div>
    <div class='panel matrix-wrap'><table class='matrix'><thead><tr><th></th>$(($overlapHeader -join ''))</tr></thead><tbody>$(($overlapRowsHtml -join ''))</tbody></table></div>
  </section>
  <section class='section'>
    <div class='section-head'><h2>经营动作</h2><p>最后一屏只回答一个问题：接下来应该怎么分资源，而不是继续堆更多数据。</p></div>
    <div class='actions'>$(($strategyBlocks -join ''))</div>
  </section>
</main>
</body>
</html>
"@

$html | Set-Content -LiteralPath $OutputPath -Encoding UTF8
