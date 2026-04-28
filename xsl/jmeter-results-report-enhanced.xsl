<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">

<xsl:output method="html" indent="yes" encoding="UTF-8" doctype-public="-//W3C//DTD HTML 4.01 Transitional//EN" />

<!-- 参数定义 -->
<xsl:param name="titleReport" select="'JMeter 性能测试报告'"/>
<xsl:param name="dateReport" select="'date not defined'"/>

<!-- 主模板 -->
<xsl:template match="testResults">
    <html>
        <head>
            <meta charset="UTF-8" />
            <title><xsl:value-of select="$titleReport" /></title>
            <script src="https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js"></script>
            <style type="text/css">
                * {
                    margin: 0;
                    padding: 0;
                    box-sizing: border-box;
                }
                
                body {
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    background-color: #f5f7fa;
                    color: #333;
                    line-height: 1.6;
                }
                
                .container {
                    max-width: 1400px;
                    margin: 0 auto;
                    padding: 20px;
                }
                
                .header {
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white;
                    padding: 30px;
                    border-radius: 8px;
                    margin-bottom: 30px;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                }
                
                .header h1 {
                    font-size: 32px;
                    margin-bottom: 10px;
                }
                
                .header p {
                    font-size: 14px;
                    opacity: 0.9;
                }
                
                .metrics {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                    gap: 15px;
                    margin-bottom: 30px;
                }
                
                .metric {
                    background: white;
                    padding: 20px;
                    border-radius: 8px;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                    text-align: center;
                    border-left: 4px solid #3498db;
                }
                
                .metric.success {
                    border-left-color: #27ae60;
                }
                
                .metric.error {
                    border-left-color: #e74c3c;
                }
                
                .metric.total {
                    border-left-color: #3498db;
                }
                
                .metric.avg {
                    border-left-color: #f39c12;
                }
                
                .metric-label {
                    font-size: 12px;
                    color: #7f8c8d;
                    text-transform: uppercase;
                    margin-bottom: 8px;
                    font-weight: 600;
                }
                
                .metric-value {
                    font-size: 36px;
                    font-weight: bold;
                    color: #2c3e50;
                }
                
                .metric-unit {
                    font-size: 12px;
                    color: #95a5a6;
                    margin-top: 5px;
                }
                
                .charts-grid {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
                    gap: 20px;
                    margin-bottom: 30px;
                }
                
                .chart-container {
                    background: white;
                    padding: 20px;
                    border-radius: 8px;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                }
                
                .chart-title {
                    font-size: 16px;
                    font-weight: bold;
                    color: #2c3e50;
                    margin-bottom: 15px;
                    padding-bottom: 10px;
                    border-bottom: 2px solid #ecf0f1;
                }
                
                .stats-table {
                    width: 100%;
                    background: white;
                    border-collapse: collapse;
                    border-radius: 8px;
                    overflow: hidden;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                    margin-bottom: 30px;
                }
                
                .stats-table th {
                    background-color: #34495e;
                    color: white;
                    padding: 12px;
                    text-align: left;
                    font-weight: 600;
                }
                
                .stats-table td {
                    padding: 12px;
                    border-bottom: 1px solid #ecf0f1;
                }
                
                .stats-table tr:last-child td {
                    border-bottom: none;
                }
                
                .stats-table tr:nth-child(even) {
                    background-color: #f8f9fa;
                }
                
                .stats-table tr:hover {
                    background-color: #f0f0f0;
                }
                
                .section-title {
                    font-size: 20px;
                    font-weight: bold;
                    color: #2c3e50;
                    margin-top: 30px;
                    margin-bottom: 20px;
                    padding-bottom: 10px;
                    border-bottom: 3px solid #667eea;
                }
                
                .info-box {
                    background: #e8f4f8;
                    border-left: 4px solid #3498db;
                    padding: 15px;
                    border-radius: 4px;
                    margin-bottom: 20px;
                }
                
                .info-box strong {
                    color: #2980b9;
                }
                
                .detail-table {
                    width: 100%;
                    background: white;
                    border-collapse: collapse;
                    border-radius: 8px;
                    overflow: hidden;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                    margin-bottom: 30px;
                }
                
                .detail-table th {
                    background-color: #34495e;
                    color: white;
                    padding: 12px;
                    text-align: left;
                    font-weight: 600;
                }
                
                .detail-table td {
                    padding: 10px 12px;
                    border-bottom: 1px solid #ecf0f1;
                }
                
                .detail-table tr:nth-child(even) {
                    background-color: #f8f9fa;
                }
                
                .detail-table tr:hover {
                    background-color: #f0f0f0;
                }
                
                .status-success {
                    color: #27ae60;
                    font-weight: bold;
                }
                
                .status-error {
                    color: #e74c3c;
                    font-weight: bold;
                }
                
                .footer {
                    text-align: center;
                    color: #7f8c8d;
                    font-size: 12px;
                    margin-top: 40px;
                    padding-top: 20px;
                    border-top: 1px solid #ecf0f1;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <!-- 页头 -->
                <div class="header">
                    <h1>📊 <xsl:value-of select="$titleReport" /></h1>
                    <p>生成时间: <xsl:value-of select="$dateReport" /></p>
                </div>
                
                <!-- 关键指标 -->
                <xsl:call-template name="metrics" />
                
                <!-- 图表区域 -->
                <xsl:call-template name="charts" />
                
                <!-- 详细统计 -->
                <xsl:call-template name="statistics" />
                
                <!-- 详细数据表 -->
                <xsl:call-template name="details" />
                
                <!-- 页脚 -->
                <div class="footer">
                    <p>此报告由 JMeter 自动生成 | 基于 XSL 样式表转换</p>
                </div>
            </div>
        </body>
    </html>
</xsl:template>

<!-- 关键指标模板 -->
<xsl:template name="metrics">
    <xsl:variable name="totalCount" select="count(/testResults/*)" />
    <xsl:variable name="successCount" select="count(/testResults/*[@s='true'])" />
    <xsl:variable name="failureCount" select="$totalCount - $successCount" />
    <xsl:variable name="successRate" select="$successCount div $totalCount * 100" />
    <xsl:variable name="totalTime" select="sum(/testResults/*/@t)" />
    <xsl:variable name="avgTime" select="$totalTime div $totalCount" />
    
    <div class="metrics">
        <div class="metric total">
            <div class="metric-label">总请求数</div>
            <div class="metric-value"><xsl:value-of select="$totalCount" /></div>
        </div>
        <div class="metric success">
            <div class="metric-label">成功请求</div>
            <div class="metric-value"><xsl:value-of select="$successCount" /></div>
            <div class="metric-unit">成功率: <xsl:value-of select="format-number($successRate, '0.00')" />%</div>
        </div>
        <div class="metric error">
            <div class="metric-label">失败请求</div>
            <div class="metric-value"><xsl:value-of select="$failureCount" /></div>
            <div class="metric-unit">失败率: <xsl:value-of select="format-number(100 - $successRate, '0.00')" />%</div>
        </div>
        <div class="metric avg">
            <div class="metric-label">平均响应时间</div>
            <div class="metric-value"><xsl:value-of select="format-number($avgTime, '0.00')" /></div>
            <div class="metric-unit">ms</div>
        </div>
    </div>
</xsl:template>

<!-- 图表模板 -->
<xsl:template name="charts">
    <xsl:variable name="totalCount" select="count(/testResults/*)" />
    <xsl:variable name="successCount" select="count(/testResults/*[@s='true'])" />
    <xsl:variable name="failureCount" select="$totalCount - $successCount" />
    
    <div class="charts-grid">
        <!-- 成功率饼图 -->
        <div class="chart-container">
            <div class="chart-title">✅ 成功率分布</div>
            <canvas id="successChart"></canvas>
        </div>
        
        <!-- 响应码分布 -->
        <div class="chart-container">
            <div class="chart-title">📡 响应码分布</div>
            <canvas id="responseCodeChart"></canvas>
        </div>
    </div>
    
    <script>
        // 成功率饼图
        new Chart(document.getElementById('successChart'), {
            type: 'doughnut',
            data: {
                labels: ['成功', '失败'],
                datasets: [{
                    data: [<xsl:value-of select="$successCount" />, <xsl:value-of select="$failureCount" />],
                    backgroundColor: ['#27ae60', '#e74c3c'],
                    borderColor: ['#229954', '#c0392b'],
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { position: 'bottom' }
                }
            }
        });
        
        // 响应码分布
        var responseCodes = {
            <xsl:for-each select="/testResults/*/@rc[not(. = preceding::*/@rc)]">
                '<xsl:value-of select="." />': <xsl:value-of select="count(/testResults/*[@rc = current()])" /><xsl:if test="position() != last()">,</xsl:if>
            </xsl:for-each>
        };
        
        new Chart(document.getElementById('responseCodeChart'), {
            type: 'bar',
            data: {
                labels: Object.keys(responseCodes),
                datasets: [{
                    label: '请求数',
                    data: Object.values(responseCodes),
                    backgroundColor: '#3498db',
                    borderColor: '#2980b9',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                scales: {
                    y: { beginAtZero: true }
                }
            }
        });
    </script>
</xsl:template>

<!-- 统计信息模板 -->
<xsl:template name="statistics">
    <xsl:variable name="totalCount" select="count(/testResults/*)" />
    <xsl:variable name="totalTime" select="sum(/testResults/*/@t)" />
    <xsl:variable name="avgTime" select="$totalTime div $totalCount" />
    <xsl:variable name="minTime">
        <xsl:call-template name="min">
            <xsl:with-param name="nodes" select="/testResults/*/@t" />
        </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="maxTime">
        <xsl:call-template name="max">
            <xsl:with-param name="nodes" select="/testResults/*/@t" />
        </xsl:call-template>
    </xsl:variable>
    
    <h2 class="section-title">📊 响应时间统计 (单位: ms)</h2>
    <table class="stats-table">
        <thead>
            <tr>
                <th>指标</th>
                <th>值</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>最小响应时间</td>
                <td><strong><xsl:value-of select="$minTime" /></strong></td>
            </tr>
            <tr>
                <td>最大响应时间</td>
                <td><strong><xsl:value-of select="$maxTime" /></strong></td>
            </tr>
            <tr>
                <td>平均响应时间</td>
                <td><strong><xsl:value-of select="format-number($avgTime, '0.00')" /></strong></td>
            </tr>
            <tr>
                <td>总请求数</td>
                <td><strong><xsl:value-of select="$totalCount" /></strong></td>
            </tr>
        </tbody>
    </table>
</xsl:template>

<!-- 详细数据表模板 -->
<xsl:template name="details">
    <h2 class="section-title">📋 详细测试数据</h2>
    <div class="info-box">
        <strong>说明:</strong> 下表显示所有测试样本的详细信息，包括请求名称、状态、响应时间等。
    </div>
    <table class="detail-table">
        <thead>
            <tr>
                <th>序号</th>
                <th>请求名称</th>
                <th>状态</th>
                <th>响应时间 (ms)</th>
                <th>响应码</th>
                <th>响应信息</th>
            </tr>
        </thead>
        <tbody>
            <xsl:for-each select="/testResults/*">
                <xsl:variable name="status" select="@s" />
                <tr>
                    <td><xsl:value-of select="position()" /></td>
                    <td><xsl:value-of select="@lb" /></td>
                    <td>
                        <xsl:choose>
                            <xsl:when test="$status = 'true'">
                                <span class="status-success">✓ 成功</span>
                            </xsl:when>
                            <xsl:otherwise>
                                <span class="status-error">✗ 失败</span>
                            </xsl:otherwise>
                        </xsl:choose>
                    </td>
                    <td><xsl:value-of select="@t" /></td>
                    <td><xsl:value-of select="@rc" /></td>
                    <td><xsl:value-of select="@rm" /></td>
                </tr>
            </xsl:for-each>
        </tbody>
    </table>
</xsl:template>

<!-- 最小值模板 -->
<xsl:template name="min">
    <xsl:param name="nodes" />
    <xsl:for-each select="$nodes">
        <xsl:sort select="." data-type="number" />
        <xsl:if test="position() = 1">
            <xsl:value-of select="." />
        </xsl:if>
    </xsl:for-each>
</xsl:template>

<!-- 最大值模板 -->
<xsl:template name="max">
    <xsl:param name="nodes" />
    <xsl:for-each select="$nodes">
        <xsl:sort select="." data-type="number" order="descending" />
        <xsl:if test="position() = 1">
            <xsl:value-of select="." />
        </xsl:if>
    </xsl:for-each>
</xsl:template>

</xsl:stylesheet>
