document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('reportForm');
    const generateBtn = document.getElementById('generateBtn');
    const btnText = generateBtn.querySelector('.btn-text');
    const loader = generateBtn.querySelector('.loader');
    const resultsSection = document.getElementById('resultsSection');
    const weekDisplay = document.getElementById('weekDisplay');
    const list10 = document.getElementById('list_10');
    const list11 = document.getElementById('list_11');
    const toast = document.getElementById('toast');

    // Modal elements
    const modal = document.getElementById('previewModal');
    const closeBtn = modal.querySelector('.close-modal');
    const previewBody = document.getElementById('previewBody');
    const previewTitle = document.getElementById('previewTitle');

    // File Input UX
    ['file_10', 'file_11', 'timesheet_file'].forEach(id => {
        const input = document.getElementById(id);
        const dropArea = document.getElementById(id === 'timesheet_file' ? 'drop-area-timesheet' : id.replace('file_', 'drop-area-'));
        const msg = dropArea.querySelector('.file-msg');

        input.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                msg.textContent = e.target.files[0].name;
                dropArea.classList.add('highlight');
            }
        });

        // Drag and drop events
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }

        ['dragenter', 'dragover'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.add('highlight'), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.remove('highlight'), false);
        });

        dropArea.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            input.files = files;
            if (files.length > 0) {
                msg.textContent = files[0].name;
                dropArea.classList.add('highlight');
            }
        });
    });

    // Toggle Advanced Config
    document.querySelectorAll('.config-toggle').forEach(toggle => {
        toggle.addEventListener('click', function () {
            this.closest('.config-section').classList.toggle('active');
        });
    });

    // Form Submit
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        btnText.textContent = 'Đang xử lý...';
        loader.style.display = 'inline-block';
        generateBtn.disabled = true;

        const formData = new FormData(form);
        try {
            const response = await fetch('/generate', { method: 'POST', body: formData });
            const text = await response.text();

            try {
                const result = JSON.parse(text);
                if (result.success) {
                    showToast(result.message, 'success');
                    displayResults(result);
                } else {
                    showToast(result.message, 'error');
                }
            } catch (e) {
                console.error("JSON Parse Error:", e);
                console.log("Raw Response:", text);
                showToast('Lỗi cấu trúc dữ liệu JSON. Kiểm tra Console để xem chi tiết.', 'error');
            }

        } catch (error) {
            console.error("Fetch Error:", error);
            showToast(`Lỗi kết nối: ${error.message}`, 'error');
        } finally {
            btnText.textContent = 'Tạo Báo Cáo';
            loader.style.display = 'none';
            generateBtn.disabled = false;
        }
    });

    function displayResults(data) {
        list10.innerHTML = '';
        list11.innerHTML = '';
        weekDisplay.textContent = `Tuần ${data.week_10} & ${data.week_11}`;

        const createReportCard = (filename, grade, week) => {
            const li = document.createElement('li');
            li.className = 'report-card';
            li.innerHTML = `
                <div class="card-info">
                    <i class="fa-solid fa-file-word icon"></i>
                    <div class="details">
                        <span class="filename" title="${filename}">${filename}</span>
                        <span class="filesize">Khối ${grade.split('_')[1]}</span>
                    </div>
                </div>
                <div class="card-actions">
                    <button class="action-btn preview-btn" data-grade="${grade}" data-week="${week}" data-file="${filename}" title="Xem trước">
                        <i class="fa-solid fa-eye"></i>
                    </button>
                    <a href="/download/${grade}/${week}/${filename}" class="action-btn download-btn" title="Tải xuống">
                        <i class="fa-solid fa-download"></i>
                    </a>
                </div>
            `;
            // Attach listener directly to the preview button
            li.querySelector('.preview-btn').onclick = function () {
                openPreview(this.dataset.grade, this.dataset.week, this.dataset.file);
            };
            return li;
        };

        data.reports_10.forEach(f => list10.appendChild(createReportCard(f, 'Grade_10', data.week_10)));
        data.reports_11.forEach(f => list11.appendChild(createReportCard(f, 'Grade_11', data.week_11)));

        // Dashboard
        if (data.stats_10 || data.stats_11) {
            displayDashboard(data.stats_10, data.stats_11);
        }

        resultsSection.style.display = 'block';
        document.getElementById('dashboardSection').style.display = 'block'; // Show dashboard
        resultsSection.scrollIntoView({ behavior: 'smooth' });

        document.getElementById('downloadZip10').onclick = () => window.location.href = `/download-zip/Grade_10/${data.week_10}`;
        document.getElementById('downloadZip11').onclick = () => window.location.href = `/download-zip/Grade_11/${data.week_11}`;
    }

    // Store chart instances globally within the closure to manage lifecycle across re-renders
    const chartInstances = {};

    function displayDashboard(stats10, stats11) {
        const dashboard = document.getElementById('dashboardSection');
        const tabsContainer = document.getElementById('dashboardTabs');
        const tabs = tabsContainer.querySelectorAll('.tab-btn');

        // Show tabs if data exists for both, or at least one
        tabsContainer.style.display = 'flex';

        // Determine initial active tab
        let activeGrade = '10';
        if (!stats10 && stats11) activeGrade = '11';

        // Tab Switching Logic
        tabs.forEach(tab => {
            tab.onclick = () => {
                tabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                activeGrade = tab.dataset.grade;
                renderDashboard(activeGrade);
            };
        });

        // Initial Render
        // Set active class on correct tab
        tabs.forEach(t => t.classList.toggle('active', t.dataset.grade === activeGrade));
        renderDashboard(activeGrade);

        function renderDashboard(grade) {
            // Destroy existing charts before removing canvas from DOM
            ['chart-efficiency', 'chart-time', 'chart-status'].forEach(id => {
                if (chartInstances[id]) {
                    chartInstances[id].destroy();
                    delete chartInstances[id];
                }
            });

            dashboard.innerHTML = '';
            const stats = grade === '10' ? stats10 : stats11;

            if (!stats) {
                dashboard.innerHTML = `<div class="glass-panel" style="padding: 20px; text-align: center; color: #64748b;">Chưa có dữ liệu phân tích cho Khối ${grade}</div>`;
                return;
            }

            // Prepare Data
            const labels = stats.class_stats.map(s => s.className);
            const timeData = stats.class_stats.map(s => s.avgTotalTime);

            // New Data for Progress Chart
            const onTrackData = stats.class_stats.map(s => s.onTrack);
            const behindData = stats.class_stats.map(s => s.behind);

            // HTML Structure
            const html = `
                <div class="dashboard-grid">
                    <!-- Chart 1: Progress Overview -->
                    <div class="card glass-panel chart-card full-width">
                        <div class="card-header">
                            <h3><i class="fa-solid fa-chart-line"></i> Số lượng học sinh theo tiến độ</h3>
                        </div>
                        <div class="chart-box-glass">
                            <canvas id="chart-efficiency"></canvas>
                        </div>
                    </div>

                    <!-- Chart 2: Time Investment -->
                    <div class="card glass-panel chart-card">
                        <div class="card-header">
                            <h3><i class="fa-solid fa-clock"></i> Thời gian học trung bình</h3>
                        </div>
                        <div class="chart-box-glass">
                            <canvas id="chart-time"></canvas>
                        </div>
                    </div>

                    <!-- Chart 3: Status Distribution -->
                    <div class="card glass-panel chart-card">
                        <div class="card-header">
                            <h3><i class="fa-solid fa-chart-pie"></i> Tỷ lệ hoàn thành</h3>
                        </div>
                        <div class="chart-box-glass">
                            <canvas id="chart-status"></canvas>
                        </div>
                    </div>
                </div>
            `;

            dashboard.innerHTML = html;

            // Common Chart Options
            const commonOptions = {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { labels: { color: '#334155', font: { family: 'Inter' } } },
                    tooltip: {
                        backgroundColor: 'rgba(255, 255, 255, 0.9)',
                        titleColor: '#0f172a',
                        bodyColor: '#334155',
                        borderColor: '#e2e8f0',
                        borderWidth: 1,
                        padding: 10,
                        displayColors: true,
                        usePointStyle: true,
                    }
                },
                scales: {
                    x: { grid: { display: false }, ticks: { color: '#64748b' } },
                    y: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b' } }
                }
            };

            // 1. Efficiency Chart (Students On Track vs Behind)
            chartInstances['chart-efficiency'] = new Chart(document.getElementById('chart-efficiency'), {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [
                        {
                            label: 'Kịp tiến độ (On Track)',
                            data: onTrackData,
                            backgroundColor: 'rgba(16, 185, 129, 0.7)', // Emerald
                            borderColor: 'rgba(16, 185, 129, 1)',
                            borderWidth: 1,
                            borderRadius: 4,
                            order: 1
                        },
                        {
                            label: 'Chậm tiến độ (Behind)',
                            data: behindData,
                            backgroundColor: 'rgba(239, 68, 68, 0.6)', // Red
                            borderColor: 'rgba(239, 68, 68, 1)',
                            borderWidth: 1,
                            borderRadius: 4,
                            order: 2
                        }
                    ]
                },
                options: {
                    ...commonOptions,
                    scales: {
                        ...commonOptions.scales,
                        y: { ...commonOptions.scales.y, title: { display: true, text: 'Số học sinh' }, stacked: true },
                        x: { ...commonOptions.scales.x, stacked: true }
                    }
                }
            });

            // 2. Time Chart (Dual Axis)
            chartInstances['chart-time'] = new Chart(document.getElementById('chart-time'), {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [
                        {
                            label: 'Tổng TG học TB (Phút/HS)',
                            data: timeData,
                            backgroundColor: 'rgba(59, 130, 246, 0.6)', // Blue
                            borderColor: 'rgba(59, 130, 246, 1)',
                            borderWidth: 1,
                            borderRadius: 4,
                            yAxisID: 'y',
                            order: 2
                        },
                        {
                            label: 'TG làm đạt 1 bài (Phút/Bài học)',
                            data: stats.class_stats.map(s => s.avgTimePerStudied),
                            type: 'line',
                            borderColor: '#f59e0b', // Amber/Orange for visibility
                            borderWidth: 2,
                            pointBackgroundColor: '#fff',
                            pointBorderColor: '#f59e0b',
                            pointRadius: 4,
                            yAxisID: 'y1',
                            order: 1
                        }
                    ]
                },
                options: {
                    ...commonOptions,
                    scales: {
                        ...commonOptions.scales,
                        y: {
                            ...commonOptions.scales.y,
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: { display: true, text: 'Tổng thời gian (Phút)' }
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            grid: { drawOnChartArea: false }, // Only draw grid for primary axis
                            title: { display: true, text: 'TG trung bình/bài (Phút)' }
                        }
                    }
                }
            });

            // 3. Status Chart (Vietnamese Labels & Percentages)
            const statusLabels = ['Vượt kế hoạch', 'Đúng kế hoạch', 'Chậm hơn kế hoạch'];
            const statusKeys = ['vượt kế hoạch', 'đúng kế hoạch', 'chậm hơn kế hoạch'];
            const statusData = statusKeys.map(key => stats.status_counts[key] || 0);
            const totalStatus = statusData.reduce((a, b) => a + b, 0);

            chartInstances['chart-status'] = new Chart(document.getElementById('chart-status'), {
                type: 'doughnut',
                data: {
                    labels: statusLabels,
                    datasets: [{
                        data: statusData,
                        backgroundColor: [
                            'rgba(16, 185, 129, 0.85)', // Emerald/Green - Vượt
                            'rgba(59, 130, 246, 0.85)',  // Blue - Đúng
                            'rgba(239, 68, 68, 0.85)'    // Rose/Red - Chậm
                        ],
                        borderWidth: 0,
                        hoverOffset: 12
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '65%',
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                usePointStyle: true,
                                padding: 20,
                                generateLabels: (chart) => {
                                    const data = chart.data;
                                    if (data.labels.length && data.datasets.length) {
                                        return data.labels.map((label, i) => {
                                            const val = data.datasets[0].data[i];
                                            const percentage = totalStatus > 0 ? ((val / totalStatus) * 100).toFixed(1) : 0;
                                            return {
                                                text: `${label} (${percentage}%)`,
                                                fillStyle: data.datasets[0].backgroundColor[i],
                                                strokeStyle: data.datasets[0].backgroundColor[i],
                                                lineWidth: 0,
                                                hidden: false,
                                                index: i
                                            };
                                        });
                                    }
                                    return [];
                                }
                            }
                        },
                        tooltip: {
                            callbacks: {
                                label: (context) => {
                                    const val = context.raw;
                                    const percentage = totalStatus > 0 ? ((val / totalStatus) * 100).toFixed(1) : 0;
                                    return ` ${context.label}: ${val} học sinh (${percentage}%)`;
                                }
                            }
                        }
                    }
                }
            });
        }
    }

    // Preview Logic
    async function openPreview(grade, week, filename) {
        previewTitle.textContent = `Xem trước: ${filename}`;
        previewBody.innerHTML = '<div class="preview-placeholder">Đang tải nội dung...</div>';
        modal.style.display = 'block';

        try {
            const response = await fetch(`/preview/${grade}/${week}/${encodeURIComponent(filename)}`);
            if (response.ok) {
                previewBody.innerHTML = await response.text();
            } else {
                previewBody.innerHTML = `<div class="error-msg">Lỗi: Không thể tải bản xem trước (HTTP ${response.status}).</div>`;
            }
        } catch (error) {
            previewBody.innerHTML = `<div class="error-msg">Lỗi kết nối server.</div>`;
        }
    }

    // Close Modal
    closeBtn.onclick = () => modal.style.display = 'none';
    window.onclick = (e) => { if (e.target == modal) modal.style.display = 'none'; };

    function showToast(message, type) {
        toast.textContent = message;
        toast.className = 'toast show';
        toast.style.backgroundColor = type === 'error' ? '#ef4444' : '#10b981';
        setTimeout(() => toast.className = 'toast', 3000);
    }
});
