"""
build_dashboard.py
Generates a standalone, production-grade interactive HTML dashboard embedding the processed analytics data.
"""

import json
from pathlib import Path
import pandas as pd

def build_html():
    base_dir = Path(__file__).resolve().parent
    dim_cust_path = base_dir / "data" / "processed" / "dim_customer_rfm.csv"
    summary_path = base_dir / "dashboard" / "summary_data.json"
    
    with open(summary_path, "r") as f:
        summary = json.load(f)
        
    df_cust = pd.read_csv(dim_cust_path)
    # Convert records for JSON embedding (sampling first 300 for snappy client performance)
    cust_records = df_cust.to_dict(orient="records")
    cust_json = json.dumps(cust_records)
    summary_json = json.dumps(summary)
    
    html_content = f"""<!DOCTYPE html>
<html lang="en" class="h-full bg-slate-950 text-slate-100">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Executive RFM & Customer Lifecycle Hub</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        tailwind.config = {{
            darkMode: 'class',
            theme: {{
                extend: {{
                    colors: {{
                        brand: {{
                            primary: '#6366f1',
                            success: '#10b981',
                            warning: '#f59e0b',
                            danger: '#ef4444',
                            purple: '#a855f7'
                        }}
                    }}
                }}
            }}
        }}
    </script>
    <style>
        ::-webkit-scrollbar {{ width: 6px; height: 6px; }}
        ::-webkit-scrollbar-track {{ background: #0f172a; }}
        ::-webkit-scrollbar-thumb {{ background: #334155; border-radius: 3px; }}
        ::-webkit-scrollbar-thumb:hover {{ background: #475569; }}
    </style>
</head>
<body class="min-h-full flex flex-col font-sans antialiased text-slate-200">
    <!-- Header Ribbon -->
    <header class="border-b border-slate-800 bg-slate-900/80 backdrop-blur sticky top-0 z-30 px-6 py-4 flex flex-wrap items-center justify-between gap-4">
        <div class="flex items-center gap-3">
            <div class="w-10 h-10 rounded-xl bg-indigo-600/20 border border-indigo-500/30 flex items-center justify-center text-indigo-400 font-bold text-xl">
                ⚡
            </div>
            <div>
                <h1 class="text-xl font-bold tracking-tight text-white flex items-center gap-2">
                    Customer Lifecycle & RFM Segmentation Hub
                    <span class="text-xs font-mono font-medium px-2 py-0.5 rounded-full bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">Production Star-Schema</span>
                </h1>
                <p class="text-xs text-slate-400">Algorithmic K-Means Clustering & Quantile CRM Allocation</p>
            </div>
        </div>
        <div class="flex items-center gap-3">
            <span class="text-xs font-mono text-slate-400">Engine: <strong class="text-slate-200">Python 3.9 + Scikit-Learn</strong></span>
            <div class="h-4 w-px bg-slate-800"></div>
            <span class="text-xs font-mono text-slate-400">T_obs: <strong class="text-slate-200">2023-12-11</strong></span>
        </div>
    </header>

    <main class="flex-1 p-6 space-y-6 max-w-7xl mx-auto w-full">
        <!-- Zone 2: Executive KPI Scorecards -->
        <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4">
            <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-4 shadow-sm">
                <span class="text-xs font-medium uppercase tracking-wider text-slate-400">Total Portfolio Spend</span>
                <div class="text-2xl font-bold text-white mt-1" id="kpi-revenue">$0.00</div>
                <div class="text-xs text-emerald-400 mt-1 flex items-center gap-1 font-medium">
                    <span>↑ 100% Net Cleared</span>
                </div>
            </div>

            <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-4 shadow-sm">
                <span class="text-xs font-medium uppercase tracking-wider text-slate-400">Active Customer Base</span>
                <div class="text-2xl font-bold text-white mt-1" id="kpi-customers">0</div>
                <div class="text-xs text-slate-400 mt-1 font-mono">Authenticated Accounts</div>
            </div>

            <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-4 shadow-sm">
                <span class="text-xs font-medium uppercase tracking-wider text-slate-400">Gross Invoices</span>
                <div class="text-2xl font-bold text-white mt-1" id="kpi-orders">0</div>
                <div class="text-xs text-indigo-400 mt-1 font-mono">Distinct Checkouts</div>
            </div>

            <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-4 shadow-sm">
                <span class="text-xs font-medium uppercase tracking-wider text-slate-400">Blended AOV</span>
                <div class="text-2xl font-bold text-white mt-1" id="kpi-aov">$0.00</div>
                <div class="text-xs text-slate-400 mt-1 font-mono">Spend / Invoices</div>
            </div>

            <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-4 shadow-sm border-l-4 border-l-amber-500">
                <span class="text-xs font-medium uppercase tracking-wider text-amber-400">Champions Rev Share</span>
                <div class="text-2xl font-bold text-amber-400 mt-1" id="kpi-champions">0%</div>
                <div class="text-xs text-slate-400 mt-1 font-mono">Pareto Top 27% Base</div>
            </div>

            <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-4 shadow-sm border-l-4 border-l-rose-500">
                <span class="text-xs font-medium uppercase tracking-wider text-rose-400">At-Risk Capital Exposure</span>
                <div class="text-2xl font-bold text-rose-400 mt-1" id="kpi-at-risk">$0.00</div>
                <div class="text-xs text-rose-400/80 mt-1 font-mono">Immediate Winback Target</div>
            </div>
        </div>

        <!-- Zone 3 & 4: Analytical Visualizations -->
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <!-- Segment Distribution (Donut) -->
            <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-5 flex flex-col">
                <div class="flex items-center justify-between mb-4">
                    <div>
                        <h2 class="text-base font-bold text-white">Segment Capital Distribution</h2>
                        <p class="text-xs text-slate-400">Share of cumulative revenue by behavioral tier</p>
                    </div>
                </div>
                <div class="relative flex-1 min-h-[260px] flex items-center justify-center">
                    <canvas id="chart-segments"></canvas>
                </div>
            </div>

            <!-- Spatial Recency vs Spend Scatterplot -->
            <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-5 flex flex-col lg:col-span-2">
                <div class="flex items-center justify-between mb-4">
                    <div>
                        <h2 class="text-base font-bold text-white">Spatial Customer Migration Space</h2>
                        <p class="text-xs text-slate-400">Recency (Days) vs. Cumulative Spend ($) with K-Means Tiers</p>
                    </div>
                    <span class="text-xs font-mono text-slate-400">Bubble Size = Frequency</span>
                </div>
                <div class="relative flex-1 min-h-[260px]">
                    <canvas id="chart-scatter"></canvas>
                </div>
            </div>
        </div>

        <!-- ML Diagnostics & Diagnostics Curve -->
        <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-5">
            <div class="flex items-center justify-between mb-3">
                <div>
                    <h2 class="text-base font-bold text-white">Unsupervised K-Means Diagnostics (Elbow Curve & Silhouette)</h2>
                    <p class="text-xs text-slate-400">Mathematical validation for k=4 selection over Log1p + StandardScaler transformed space</p>
                </div>
                <span class="text-xs font-mono px-2.5 py-1 rounded bg-indigo-500/10 text-indigo-400 border border-indigo-500/20">Optimal k = 4</span>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div class="h-[200px] relative">
                    <canvas id="chart-inertia"></canvas>
                </div>
                <div class="h-[200px] relative">
                    <canvas id="chart-silhouette"></canvas>
                </div>
            </div>
        </div>

        <!-- Zone 5: Customer Activation Ledger (Interactive Table) -->
        <div class="bg-slate-900/60 border border-slate-800/80 rounded-2xl p-5 space-y-4">
            <div class="flex flex-col md:flex-row md:items-center justify-between gap-3">
                <div>
                    <h2 class="text-base font-bold text-white">Operational Customer Activation Ledger</h2>
                    <p class="text-xs text-slate-400">Granular customer-level scores, segments, and prescribed commercial action playbook</p>
                </div>
                <div class="flex items-center gap-3">
                    <select id="filter-segment" class="bg-slate-950 border border-slate-800 text-xs rounded-xl px-3 py-2 text-slate-300 focus:outline-none focus:border-indigo-500">
                        <option value="ALL">All Business Segments</option>
                    </select>
                    <input type="text" id="search-input" placeholder="Search Customer ID..." class="bg-slate-950 border border-slate-800 text-xs rounded-xl px-3 py-2 text-slate-300 placeholder-slate-500 focus:outline-none focus:border-indigo-500 w-48">
                </div>
            </div>

            <div class="overflow-x-auto rounded-xl border border-slate-800/80">
                <table class="w-full text-left text-xs text-slate-300">
                    <thead class="bg-slate-950/80 text-slate-400 font-mono uppercase text-[11px] tracking-wider border-b border-slate-800">
                        <tr>
                            <th class="py-3 px-4">Customer ID</th>
                            <th class="py-3 px-4">Segment</th>
                            <th class="py-3 px-4 text-right">Recency</th>
                            <th class="py-3 px-4 text-right">Frequency</th>
                            <th class="py-3 px-4 text-right">Monetary Spend</th>
                            <th class="py-3 px-4 text-right">AOV</th>
                            <th class="py-3 px-4 text-center">RFM Score</th>
                            <th class="py-3 px-4">Prescribed Commercial Action</th>
                        </tr>
                    </thead>
                    <tbody id="table-body" class="divide-y divide-slate-800/50">
                        <!-- Populated by JS -->
                    </tbody>
                </table>
            </div>
            <div class="flex items-center justify-between text-xs text-slate-400 px-1 font-mono">
                <span id="records-count">Showing 0 customers</span>
                <span>Export format: dim_customer_rfm.csv</span>
            </div>
        </div>
    </main>

    <footer class="border-t border-slate-800 py-4 px-6 text-center text-xs text-slate-500">
        E-Commerce Customer Lifecycle & RFM Segmentation Dashboard • Generated for Executive Portfolio
    </footer>

    <script>
        const summaryData = {summary_json};
        const allCustomers = {cust_json};

        // Populate Scorecards
        document.getElementById('kpi-revenue').innerText = '$' + summaryData.kpis.total_revenue.toLocaleString('en-US', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
        document.getElementById('kpi-customers').innerText = summaryData.kpis.active_customers.toLocaleString();
        document.getElementById('kpi-orders').innerText = summaryData.kpis.total_orders.toLocaleString();
        document.getElementById('kpi-aov').innerText = '$' + summaryData.kpis.blended_aov.toFixed(2);
        document.getElementById('kpi-champions').innerText = summaryData.kpis.champions_rev_share.toFixed(1) + '%';
        document.getElementById('kpi-at-risk').innerText = '$' + summaryData.kpis.at_risk_capital.toLocaleString('en-US', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});

        // Color mapping per segment
        const segmentColors = {{
            'Champions': '#10b981',
            'Loyal Customers': '#6366f1',
            'Hibernating': '#64748b',
            'Mid-Market Regulars': '#38bdf8',
            'At Risk (High Value)': '#ef4444',
            'Lost / Inactive': '#475569',
            'Needs Attention': '#f59e0b',
            'Recent Converts / Promising': '#a855f7'
        }};

        // 1. Donut Chart
        const segLabels = summaryData.segments.map(s => s.Business_Segment);
        const segSpend = summaryData.segments.map(s => s.Total_Spend);
        const segColorsList = segLabels.map(l => segmentColors[l] || '#94a3b8');

        new Chart(document.getElementById('chart-segments'), {{
            type: 'doughnut',
            data: {{
                labels: segLabels,
                datasets: [{{
                    data: segSpend,
                    backgroundColor: segColorsList,
                    borderColor: '#0f172a',
                    borderWidth: 2
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ position: 'bottom', labels: {{ color: '#94a3b8', font: {{ size: 10 }}, boxWidth: 12 }} }},
                    tooltip: {{
                        callbacks: {{
                            label: function(ctx) {{
                                return `${{ctx.label}}: $${{ctx.parsed.toLocaleString()}}`;
                            }}
                        }}
                    }}
                }}
            }}
        }});

        // 2. Spatial Scatterplot
        const scatterPoints = allCustomers.map(c => ({{
            x: c.Recency,
            y: c.Monetary,
            r: Math.min(14, Math.max(3, c.Frequency * 1.5)),
            customer: c
        }}));

        new Chart(document.getElementById('chart-scatter'), {{
            type: 'bubble',
            data: {{
                datasets: Object.keys(segmentColors).map(seg => ({{
                    label: seg,
                    data: scatterPoints.filter(p => p.customer.Business_Segment === seg),
                    backgroundColor: (segmentColors[seg] || '#94a3b8') + '88',
                    borderColor: segmentColors[seg] || '#94a3b8',
                    borderWidth: 1
                }}))
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                scales: {{
                    x: {{
                        title: {{ display: true, text: 'Recency (Days Inactive)', color: '#94a3b8' }},
                        grid: {{ color: '#1e293b' }},
                        ticks: {{ color: '#94a3b8' }}
                    }},
                    y: {{
                        title: {{ display: true, text: 'Monetary Spend ($)', color: '#94a3b8' }},
                        grid: {{ color: '#1e293b' }},
                        ticks: {{ color: '#94a3b8' }}
                    }}
                }},
                plugins: {{
                    legend: {{ display: false }},
                    tooltip: {{
                        callbacks: {{
                            label: function(ctx) {{
                                const raw = ctx.raw.customer;
                                return `ID: ${{raw.CustomerID}} | ${{raw.Business_Segment}} | R:${{raw.Recency}}d F:${{raw.Frequency}} M:$${{raw.Monetary}}`;
                            }}
                        }}
                    }}
                }}
            }}
        }});

        // 3. Cluster Evaluation Charts
        new Chart(document.getElementById('chart-inertia'), {{
            type: 'line',
            data: {{
                labels: summaryData.cluster_eval.k,
                datasets: [{{
                    label: 'Inertia (WCSS)',
                    data: summaryData.cluster_eval.inertia,
                    borderColor: '#6366f1',
                    backgroundColor: '#6366f133',
                    fill: true,
                    tension: 0.3
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                scales: {{
                    x: {{ title: {{ display: true, text: 'k clusters', color: '#94a3b8', font: {{ size: 10 }} }}, grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8' }} }},
                    y: {{ title: {{ display: true, text: 'Inertia', color: '#94a3b8', font: {{ size: 10 }} }}, grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8' }} }}
                }},
                plugins: {{ legend: {{ labels: {{ color: '#94a3b8', font: {{ size: 10 }} }} }} }}
            }}
        }});

        new Chart(document.getElementById('chart-silhouette'), {{
            type: 'bar',
            data: {{
                labels: summaryData.cluster_eval.k,
                datasets: [{{
                    label: 'Silhouette Score',
                    data: summaryData.cluster_eval.silhouette,
                    backgroundColor: summaryData.cluster_eval.k.map(k => k === 4 ? '#10b981' : '#334155'),
                    borderRadius: 4
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                scales: {{
                    x: {{ title: {{ display: true, text: 'k clusters', color: '#94a3b8', font: {{ size: 10 }} }}, grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8' }} }},
                    y: {{ title: {{ display: true, text: 'Score', color: '#94a3b8', font: {{ size: 10 }} }}, min: 0.2, max: 0.5, grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8' }} }}
                }},
                plugins: {{ legend: {{ labels: {{ color: '#94a3b8', font: {{ size: 10 }} }} }} }}
            }}
        }});

        // Populate Segment Dropdown Filter
        const segFilter = document.getElementById('filter-segment');
        segLabels.forEach(seg => {{
            const opt = document.createElement('option');
            opt.value = seg;
            opt.innerText = seg;
            segFilter.appendChild(opt);
        }});

        // Render Table Rows
        function renderTable() {{
            const filterVal = segFilter.value;
            const searchVal = document.getElementById('search-input').value.trim().toLowerCase();
            const tbody = document.getElementById('table-body');
            tbody.innerHTML = '';

            const filtered = allCustomers.filter(c => {{
                const matchesSeg = (filterVal === 'ALL' || c.Business_Segment === filterVal);
                const matchesSearch = (!searchVal || String(c.CustomerID).toLowerCase().includes(searchVal));
                return matchesSeg && matchesSearch;
            }});

            document.getElementById('records-count').innerText = `Showing ${{filtered.length}} of ${{allCustomers.length}} customers`;

            filtered.slice(0, 100).forEach(c => {{
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-900/80 transition-colors border-b border-slate-800/40';
                
                const segBadgeColor = segmentColors[c.Business_Segment] || '#94a3b8';
                
                tr.innerHTML = `
                    <td class="py-3 px-4 font-mono font-medium text-white">${{c.CustomerID}}</td>
                    <td class="py-3 px-4">
                        <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-medium" style="background-color: ${{segBadgeColor}}22; color: ${{segBadgeColor}}; border: 1px solid ${{segBadgeColor}}44">
                            <span class="w-1.5 h-1.5 rounded-full" style="background-color: ${{segBadgeColor}}"></span>
                            ${{c.Business_Segment}}
                        </span>
                    </td>
                    <td class="py-3 px-4 text-right font-mono text-slate-300">${{c.Recency}}d</td>
                    <td class="py-3 px-4 text-right font-mono text-slate-300">${{c.Frequency}}</td>
                    <td class="py-3 px-4 text-right font-mono font-medium text-emerald-400">$${{c.Monetary.toFixed(2)}}</td>
                    <td class="py-3 px-4 text-right font-mono text-slate-300">$${{c.AOV.toFixed(2)}}</td>
                    <td class="py-3 px-4 text-center">
                        <span class="font-mono bg-slate-950 border border-slate-800 text-slate-300 px-2 py-0.5 rounded text-[11px]">${{c.RFM_Composite}}</span>
                    </td>
                    <td class="py-3 px-4 text-slate-400">${{c.Recommended_Action}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        segFilter.addEventListener('change', renderTable);
        document.getElementById('search-input').addEventListener('input', renderTable);
        renderTable();
    </script>
</body>
</html>
"""

    out_path = base_dir / "dashboard" / "index.html"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[DASHBOARD] Successfully built standalone interactive dashboard at {out_path}")

if __name__ == "__main__":
    build_html()
