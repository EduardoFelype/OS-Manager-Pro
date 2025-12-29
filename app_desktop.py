from flask import Flask, render_template_string, request, jsonify, send_file
import pandas as pd
import sqlite3
from datetime import datetime
import os
from werkzeug.utils import secure_filename
import json
import re
from io import BytesIO

# --------------------
# Configuração
# --------------------
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --------------------
# HTML e CSS (Tailwind CSS via CDN para um design moderno)
# --------------------
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Ordens de Serviço - Dashboard Executivo</title>
    <!-- Tailwind CSS CDN para um design moderno e responsivo -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Chart.js para gráficos -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js"></script>
    <!-- Icones Heroicons -->
    <style>
        /* Estilos customizados para a barra lateral e transições */
        .sidebar {
            transition: width 0.3s ease;
        }
        .content {
            transition: margin-left 0.3s ease;
        }
        .page {
            display: none;
        }
        .page.active {
            display: block;
        }
        /* Estilo para a tabela com zebra e hover */
        .table-auto-scroll {
            overflow-x: auto;
        }
        .table-auto-scroll table {
            min-width: 1000px; /* Garante que a tabela não fique muito estreita */
        }
        .table-auto-scroll th, .table-auto-scroll td {
            white-space: nowrap;
        }
        /* Cores customizadas para o tema */
        .bg-primary { background-color: #1e3a8a; } /* Azul Escuro */
        .text-primary { color: #1e3a8a; }
        .hover-bg-primary-light:hover { background-color: #3b82f6; } /* Azul Claro */
    </style>
</head>
<body class="bg-gray-50">
    <div class="flex min-h-screen">
        <!-- Sidebar -->
        <div id="sidebar" class="sidebar w-64 bg-primary text-white p-4 fixed h-full shadow-2xl z-10">
            <h2 class="text-2xl font-bold mb-8 border-b border-blue-700 pb-3">OS Manager Pro</h2>
            <nav class="space-y-2">
                <button class="nav-btn w-full flex items-center space-x-3 p-3 rounded-lg text-left font-medium hover-bg-primary-light active" onclick="showPage('dashboard', this)">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h18M3 14h18m-9-4v8m-7-8v8m14-8v8"></path></svg>
                    <span>Dashboard</span>
                </button>
                <button class="nav-btn w-full flex items-center space-x-3 p-3 rounded-lg text-left font-medium hover-bg-primary-light" onclick="showPage('consulta', this)">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
                    <span>Consulta Avançada</span>
                </button>
                <button class="nav-btn w-full flex items-center space-x-3 p-3 rounded-lg text-left font-medium hover-bg-primary-light" onclick="showPage('upload', this)">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"></path></svg>
                    <span>Atualizar Dados</span>
                </button>
                <button class="nav-btn w-full flex items-center space-x-3 p-3 rounded-lg text-left font-medium hover-bg-primary-light" onclick="showPage('relatorios', this)">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>
                    <span>Relatórios</span>
                </button>
                <button class="nav-btn w-full flex items-center space-x-3 p-3 rounded-lg text-left font-medium hover-bg-primary-light" onclick="showPage('config', this)">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"></path></svg>
                    <span>Configurações</span>
                </button>
            </nav>
        </div>
        
        <!-- Conteúdo Principal -->
        <div id="content" class="content ml-64 p-8 w-full">
            <header class="mb-8 pb-4 border-b border-gray-200">
                <h1 class="text-3xl font-extrabold text-gray-800">Dashboard de Ordens de Serviço</h1>
                <p class="text-gray-500 italic">Visão executiva e consulta avançada para gestão de performance.</p>
            </header>

            <!-- Dashboard Page -->
            <div id="dashboard" class="page active">
                <section class="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-6 mb-8" id="metrics">
                    <!-- Cards de Métricas serão injetados aqui -->
                </section>

                <section class="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
                    <div class="bg-white p-6 rounded-xl shadow-lg">
                        <h3 class="text-xl font-semibold text-gray-700 mb-4">Status das Ordens</h3>
                        <canvas id="statusChart"></canvas>
                    </div>
                    <div class="bg-white p-6 rounded-xl shadow-lg">
                        <h3 class="text-xl font-semibold text-gray-700 mb-4">Valor por Status de Cotação</h3>
                        <canvas id="cotacaoChart"></canvas>
                    </div>
                </section>

                <section class="bg-white p-6 rounded-xl shadow-lg mb-8">
                    <h3 class="text-xl font-semibold text-gray-700 mb-4">Timeline de Criação das Ordens (Mensal)</h3>
                    <canvas id="timelineChart"></canvas>
                </section>

                <section class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    <div class="bg-white p-6 rounded-xl shadow-lg">
                        <h3 class="text-xl font-semibold text-gray-700 mb-4">Top 10 Clientes</h3>
                        <div id="topClientes" class="divide-y divide-gray-100"></div>
                    </div>
                    <div class="bg-white p-6 rounded-xl shadow-lg">
                        <h3 class="text-xl font-semibold text-gray-700 mb-4">Top 10 Produtos</h3>
                        <div id="topProdutos" class="divide-y divide-gray-100"></div>
                    </div>
                </section>
            </div>
            
            <!-- Consulta Page -->
            <div id="consulta" class="page">
                <h2 class="text-2xl font-bold text-gray-800 mb-6">Consulta Detalhada de Dados</h2>
                
                <!-- Filtros Avançados -->
                <div class="bg-white p-6 rounded-xl shadow-lg mb-6">
                    <h3 class="text-lg font-semibold text-gray-700 mb-4">Filtros de Busca</h3>
                    <div class="grid grid-cols-1 md:grid-cols-5 gap-4">
                        <div class="md:col-span-2">
                            <label for="searchInput" class="block text-sm font-medium text-gray-700">Cliente/Descrição</label>
                            <input type="text" id="searchInput" placeholder="Buscar por texto..." class="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 p-2 border">
                        </div>
                        <div>
                            <label for="statusFilter" class="block text-sm font-medium text-gray-700">Status da OS</label>
                            <select id="statusFilter" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 p-2 border bg-white">
                                <option value="">Todos os Status</option>
                                <!-- Opções serão carregadas via JS -->
                            </select>
                        </div>
                        <div>
                            <label for="cotacaoFilter" class="block text-sm font-medium text-gray-700">Status Cotação</label>
                            <select id="cotacaoFilter" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 p-2 border bg-white">
                                <option value="">Todos os Status</option>
                                <!-- Opções serão carregadas via JS -->
                            </select>
                        </div>
                        <div class="flex items-end space-x-2">
                            <button class="w-1/2 bg-blue-600 text-white p-2 rounded-md hover:bg-blue-700 transition duration-150 font-semibold" onclick="buscarDados()">
                                <svg class="w-5 h-5 inline-block" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
                            </button>
                            <button class="w-1/2 bg-green-600 text-white p-2 rounded-md hover:bg-green-700 transition duration-150 font-semibold" onclick="exportarDados()">
                                <svg class="w-5 h-5 inline-block" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
                            </button>
                        </div>
                    </div>
                </div>

                <!-- Resultados da Consulta -->
                <div class="bg-white p-6 rounded-xl shadow-lg">
                    <h3 id="resultadosTitle" class="text-lg font-semibold text-gray-700 mb-4">Resultados: 0 registros</h3>
                    <div class="table-auto-scroll">
                        <table class="min-w-full divide-y divide-gray-200">
                            <thead class="bg-gray-50">
                                <tr>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">ID</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Descrição</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status Cotação</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Produto</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Cliente</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Valor (R$)</th>
                                </tr>
                            </thead>
                            <tbody id="resultadosBody" class="bg-white divide-y divide-gray-200">
                                <!-- Linhas de resultado serão injetadas aqui -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <!-- Upload Page -->
            <div id="upload" class="page">
                <h2 class="text-2xl font-bold text-gray-800 mb-6">Atualização de Dados</h2>
                
                <div class="bg-blue-50 border-l-4 border-blue-400 text-blue-700 p-4 mb-6" role="alert">
                    <p class="font-bold">Instruções Importantes:</p>
                    <p class="text-sm">1. Faça upload da planilha <code>CARGA_PAINEL.xlsx</code> atualizada.<br>2. Escolha entre **Substituir** (apaga tudo e insere o novo) ou **Adicionar** (mantém o existente e insere o novo).</p>
                </div>
                
                <div class="bg-white p-6 rounded-xl shadow-lg">
                    <h3 class="text-xl font-semibold text-gray-700 mb-4">Selecione a planilha (.xlsx ou .xls)</h3>
                    <input type="file" id="fileInput" accept=".xlsx,.xls" class="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100">
                    
                    <div class="mt-6 flex space-x-4">
                        <button class="btn-upload bg-red-600 text-white p-3 rounded-lg hover:bg-red-700 transition duration-150 font-semibold flex-1" onclick="uploadFile(true)">
                            <svg class="w-5 h-5 inline-block mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                            Substituir Dados Existentes
                        </button>
                        <button class="btn-upload bg-green-600 text-white p-3 rounded-lg hover:bg-green-700 transition duration-150 font-semibold flex-1" onclick="uploadFile(false)">
                            <svg class="w-5 h-5 inline-block mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 6v6m0 0v6m0-6h6m-6 0H6"></path></svg>
                            Adicionar aos Dados Existentes
                        </button>
                    </div>
                    
                    <div id="uploadAlert" class="mt-6"></div>
                    <div class="loading hidden text-center mt-6" id="uploadLoading">
                        <div class="animate-spin rounded-full h-10 w-10 border-b-2 border-blue-600 mx-auto"></div>
                        <p class="text-gray-500 mt-2">Processando, aguarde...</p>
                    </div>
                </div>
            </div>
            
            <!-- Relatórios Page -->
            <div id="relatorios" class="page">
                <h2 class="text-2xl font-bold text-gray-800 mb-6">Relatórios Gerenciais</h2>
                
                <section class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8" id="relatoriosMetrics">
                    <!-- Cards de Métricas serão injetados aqui -->
                </section>
                
                <div class="bg-white p-6 rounded-xl shadow-lg">
                    <h3 class="text-xl font-semibold text-gray-700 mb-4">Performance Detalhada por Status</h3>
                    <div class="table-auto-scroll">
                        <table class="min-w-full divide-y divide-gray-200">
                            <thead class="bg-gray-50">
                                <tr>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status OS</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status Cotação</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Quantidade</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Total (R$)</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Ticket Médio (R$)</th>
                                </tr>
                            </thead>
                            <tbody id="performanceBody" class="bg-white divide-y divide-gray-200">
                                <!-- Linhas de performance serão injetadas aqui -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <!-- Config Page -->
            <div id="config" class="page">
                <h2 class="text-2xl font-bold text-gray-800 mb-6">Configurações do Sistema</h2>
                
                <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8" id="configMetrics">
                    <!-- Cards de Métricas de Configuração serão injetados aqui -->
                </div>
                
                <div class="bg-white p-6 rounded-xl shadow-lg mb-6">
                    <h3 class="text-xl font-semibold text-gray-700 mb-4">Estrutura da Planilha Esperada</h3>
                    <p class="text-sm text-gray-500 mb-3">O sistema busca as seguintes colunas na planilha de upload:</p>
                    <div id="colunasEsperadas" class="bg-gray-50 p-4 rounded-lg text-sm font-mono text-gray-700"></div>
                </div>
                
                <div class="bg-red-50 border-l-4 border-red-400 text-red-700 p-4 mb-6" role="alert">
                    <h3 class="text-xl font-semibold text-red-800 mb-4">Limpeza de Dados</h3>
                    <p class="text-sm mb-4"><strong>ATENÇÃO:</strong> Esta ação é **irreversível** e apagará todos os registros do banco de dados.</p>
                    <button class="bg-red-600 text-white p-3 rounded-lg hover:bg-red-700 transition duration-150 font-semibold" onclick="limparDados()">
                        <svg class="w-5 h-5 inline-block mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                        Limpar Todos os Dados
                    </button>
                </div>
            </div>
        </div>
    </div>

    <script>
        let statusChart, cotacaoChart, timelineChart;
        
        // Função utilitária para formatar valores monetários
        const formatCurrency = (value) => {
            return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);
        };

        // Função utilitária para obter a cor do status
        const getStatusColor = (status) => {
            const statusMap = {
                'Concluído': 'bg-green-100 text-green-800',
                'Pendente': 'bg-yellow-100 text-yellow-800',
                'Aberto': 'bg-blue-100 text-blue-800',
                'Liberado': 'bg-indigo-100 text-indigo-800',
                'Aprovado': 'bg-purple-100 text-purple-800',
                'Em Andamento': 'bg-orange-100 text-orange-800',
                'Não Definido': 'bg-gray-100 text-gray-800',
                'Sem Status': 'bg-gray-100 text-gray-800',
            };
            return statusMap[status] || 'bg-gray-100 text-gray-800';
        };

        function showPage(pageName, buttonElement) {
            console.log('Mostrando página:', pageName);
            
            document.querySelectorAll('.nav-btn').forEach(btn => {
                btn.classList.remove('bg-blue-700', 'bg-blue-800');
            });
            
            if (buttonElement) {
                buttonElement.classList.add('bg-blue-700', 'bg-blue-800');
            }
            
            document.querySelectorAll('.page').forEach(page => {
                page.classList.remove('active');
            });
            
            document.getElementById(pageName).classList.add('active');
            
            // Carregar dados específicos da página
            if (pageName === 'dashboard') loadDashboard();
            else if (pageName === 'consulta') {
                loadFilterOptions();
                buscarDados();
            }
            else if (pageName === 'relatorios') loadRelatorios();
            else if (pageName === 'config') loadConfig();
        }
        
        function loadFilterOptions() {
            fetch('/api/filtros')
                .then(r => r.json())
                .then(data => {
                    const statusSelect = document.getElementById('statusFilter');
                    const cotacaoSelect = document.getElementById('cotacaoFilter');
                    
                    // Limpar e adicionar opção padrão
                    statusSelect.innerHTML = '<option value="">Todos os Status</option>';
                    cotacaoSelect.innerHTML = '<option value="">Todos os Status</option>';
                    
                    data.status.forEach(s => {
                        statusSelect.innerHTML += `<option value="${s}">${s}</option>`;
                    });
                    data.status_cotacao.forEach(s => {
                        cotacaoSelect.innerHTML += `<option value="${s}">${s}</option>`;
                    });
                });
        }

        function loadDashboard() {
            if (statusChart) statusChart.destroy();
            if (cotacaoChart) cotacaoChart.destroy();
            if (timelineChart) timelineChart.destroy();
            
            fetch('/api/dashboard')
                .then(r => r.json())
                .then(data => {
                    // 1. Métricas
                    const metrics = [
                        { title: 'Total de Ordens', value: data.metricas.total, icon: '<svg class="w-6 h-6 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"></path></svg>' },
                        { title: 'Concluídas', value: data.metricas.concluidas, icon: '<svg class="w-6 h-6 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>' },
                        { title: 'Pendentes', value: data.metricas.pendentes, icon: '<svg class="w-6 h-6 text-yellow-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>' },
                        { title: 'Valor Total', value: data.metricas.valor_total, icon: '<svg class="w-6 h-6 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 9V7a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2m2 4h10a2 2 0 002-2v-6a2 2 0 00-2-2H9a2 2 0 00-2 2v6a2 2 0 002 2zm7-5a2 2 0 11-4 0 2 2 0 014 0z"></path></svg>' },
                        { title: 'Última Atualização', value: data.metricas.ultima_atualizacao, icon: '<svg class="w-6 h-6 text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"></path></svg>' }
                    ];
                    
                    let htmlMetrics = metrics.map(m => `
                        <div class="bg-white p-5 rounded-xl shadow-md border-l-4 border-blue-600">
                            <div class="flex items-center">
                                ${m.icon}
                                <div class="ml-4">
                                    <h3 class="text-sm font-medium text-gray-500">${m.title}</h3>
                                    <p class="text-2xl font-bold text-gray-900">${m.value}</p>
                                </div>
                            </div>
                        </div>
                    `).join('');
                    document.getElementById('metrics').innerHTML = htmlMetrics;

                    // 2. Gráfico de Status
                    const statusColors = ['#3b82f6', '#ef4444', '#10b981', '#6366f1', '#f59e0b', '#4b5563'];
                    const statusData = data.status_chart.labels.map((label, index) => ({
                        label: label,
                        value: data.status_chart.values[index],
                        color: statusColors[index % statusColors.length]
                    }));

                    const ctxStatus = document.getElementById('statusChart').getContext('2d');
                    statusChart = new Chart(ctxStatus, {
                        type: 'doughnut',
                        data: {
                            labels: statusData.map(d => d.label),
                            datasets: [{
                                data: statusData.map(d => d.value),
                                backgroundColor: statusData.map(d => d.color),
                                hoverOffset: 4
                            }]
                        },
                        options: {
                            responsive: true,
                            plugins: {
                                legend: { position: 'right' },
                                title: { display: false }
                            }
                        }
                    });
                    
                    // 3. Gráfico de Cotação
                    const ctxCotacao = document.getElementById('cotacaoChart').getContext('2d');
                    cotacaoChart = new Chart(ctxCotacao, {
                        type: 'bar',
                        data: {
                            labels: data.cotacao_chart.labels,
                            datasets: [{
                                label: 'Valor (R$)',
                                data: data.cotacao_chart.values,
                                backgroundColor: '#1e3a8a',
                                borderColor: '#1e3a8a',
                                borderWidth: 1
                            }]
                        },
                        options: {
                            responsive: true,
                            scales: {
                                y: { beginAtZero: true, ticks: { callback: (value) => formatCurrency(value) } }
                            },
                            plugins: {
                                legend: { display: false },
                                tooltip: { callbacks: { label: (context) => `Valor: ${formatCurrency(context.parsed.y)}` } }
                            }
                        }
                    });
                    
                    // 4. Gráfico de Timeline
                    const ctxTimeline = document.getElementById('timelineChart').getContext('2d');
                    timelineChart = new Chart(ctxTimeline, {
                        type: 'line',
                        data: {
                            labels: data.timeline_chart.labels,
                            datasets: [{
                                label: 'Ordens Criadas',
                                data: data.timeline_chart.values,
                                borderColor: '#3b82f6',
                                backgroundColor: 'rgba(59, 130, 246, 0.2)',
                                tension: 0.4,
                                fill: true
                            }]
                        },
                        options: {
                            responsive: true,
                            scales: {
                                y: { beginAtZero: true }
                            },
                            plugins: {
                                legend: { position: 'top' }
                            }
                        }
                    });
                    
                    // 5. Top Clientes
                    let htmlClientes = data.top_clientes.map(c => `
                        <div class="flex justify-between items-center py-2">
                            <span class="text-gray-700">${c.nome}</span>
                            <span class="px-3 py-1 text-xs font-semibold rounded-full bg-blue-100 text-blue-800">${c.count} OS</span>
                        </div>
                    `).join('');
                    document.getElementById('topClientes').innerHTML = htmlClientes || '<p class="text-gray-500 py-2">Nenhum cliente encontrado.</p>';
                    
                    // 6. Top Produtos
                    let htmlProdutos = data.top_produtos.map(p => `
                        <div class="flex justify-between items-center py-2">
                            <span class="text-gray-700">${p.nome}</span>
                            <span class="px-3 py-1 text-xs font-semibold rounded-full bg-green-100 text-green-800">${p.count} OS</span>
                        </div>
                    `).join('');
                    document.getElementById('topProdutos').innerHTML = htmlProdutos || '<p class="text-gray-500 py-2">Nenhum produto encontrado.</p>';
                });
        }
        
        function uploadFile(atualizar) {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];
            
            if (!file) {
                alert('Selecione um arquivo!');
                return;
            }
            
            const formData = new FormData();
            formData.append('file', file);
            formData.append('atualizar', atualizar);
            
            document.getElementById('uploadLoading').classList.remove('hidden');
            document.getElementById('uploadAlert').innerHTML = '';
            document.querySelectorAll('.btn-upload').forEach(btn => btn.disabled = true);
            
            fetch('/api/upload', { method: 'POST', body: formData })
                .then(r => r.json())
                .then(data => {
                    document.getElementById('uploadLoading').classList.add('hidden');
                    document.querySelectorAll('.btn-upload').forEach(btn => btn.disabled = false);
                    
                    const alertClass = data.success ? 'bg-green-100 border-green-500 text-green-700' : 'bg-red-100 border-red-500 text-red-700';
                    document.getElementById('uploadAlert').innerHTML = `
                        <div class="border-l-4 p-4 ${alertClass}" role="alert">
                            <p class="font-bold">${data.success ? 'Sucesso!' : 'Erro!'}</p>
                            <p>${data.message}</p>
                        </div>
                    `;
                    
                    if (data.success) {
                        fileInput.value = '';
                        setTimeout(() => {
                            showPage('dashboard', document.querySelector('.nav-btn')); // Volta para o dashboard
                        }, 2000);
                    }
                })
                .catch(error => {
                    document.getElementById('uploadLoading').classList.add('hidden');
                    document.querySelectorAll('.btn-upload').forEach(btn => btn.disabled = false);
                    document.getElementById('uploadAlert').innerHTML = `
                        <div class="border-l-4 p-4 bg-red-100 border-red-500 text-red-700" role="alert">
                            <p class="font-bold">Erro de Conexão!</p>
                            <p>Não foi possível comunicar com o servidor: ${error.message}</p>
                        </div>
                    `;
                });
        }
        
        function buscarDados() {
            const busca = document.getElementById('searchInput').value;
            const status = document.getElementById('statusFilter').value;
            const cotacao = document.getElementById('cotacaoFilter').value;
            
            const params = new URLSearchParams();
            if (busca) params.append('busca', busca);
            if (status) params.append('status', status);
            if (cotacao) params.append('status_cotacao', cotacao);
            
            fetch('/api/consultar?' + params.toString())
                .then(r => r.json())
                .then(data => {
                    document.getElementById('resultadosTitle').textContent = `Resultados: ${data.total} registros (Limitado a 1000)`;
                    
                    let html = data.resultados.map(r => {
                        const statusClass = getStatusColor(r.status || 'Sem Status');
                        return `
                            <tr class="hover:bg-gray-50">
                                <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${r.id}</td>
                                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${r.descricao_operacao || '-'}</td>
                                <td class="px-6 py-4 whitespace-nowrap">
                                    <span class="px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${statusClass}">
                                        ${r.status || 'Sem Status'}
                                    </span>
                                </td>
                                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${r.status_cotacao || '-'}</td>
                                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${r.denominacao_produto || '-'}</td>
                                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${r.nome_emissor_ordem || '-'}</td>
                                <td class="px-6 py-4 whitespace-nowrap text-sm font-bold text-gray-900">${formatCurrency(r.valor_pedido_bruto || 0)}</td>
                            </tr>
                        `;
                    }).join('');
                    
                    document.getElementById('resultadosBody').innerHTML = html || `
                        <tr>
                            <td colspan="7" class="px-6 py-4 text-center text-gray-500">Nenhum registro encontrado com os filtros aplicados.</td>
                        </tr>
                    `;
                });
        }
        
        function exportarDados() {
            const busca = document.getElementById('searchInput').value;
            const status = document.getElementById('statusFilter').value;
            const cotacao = document.getElementById('cotacaoFilter').value;
            
            const params = new URLSearchParams();
            if (busca) params.append('busca', busca);
            if (status) params.append('status', status);
            if (cotacao) params.append('status_cotacao', cotacao);
            
            // Redireciona para a rota de exportação com os parâmetros de filtro
            window.location.href = '/api/exportar?' + params.toString();
        }
        
        function loadRelatorios() {
            fetch('/api/relatorios')
                .then(r => r.json())
                .then(data => {
                    // Métricas
                    const metrics = [
                        { title: 'Ordens no Período', value: data.metricas.total, icon: '<svg class="w-6 h-6 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>' },
                        { title: 'Valor Total (R$)', value: formatCurrency(data.metricas.valor_total), icon: '<svg class="w-6 h-6 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V3m0 13v-5m0 0h2.879a1 1 0 01.707.293l3.424 3.424a1 1 0 01.293.707V19a2 2 0 01-2 2H5a2 2 0 01-2-2v-5.586a1 1 0 01.293-.707l3.424-3.424A1 1 0 019.121 11H12z"></path></svg>' },
                        { title: 'Ticket Médio (R$)', value: formatCurrency(data.metricas.ticket_medio), icon: '<svg class="w-6 h-6 text-purple-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 3.055A9.001 9.001 0 1020.945 13H11V3.055z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.488 9H15V3.512A9.025 9.025 0 0120.488 9z"></path></svg>' }
                    ];
                    
                    let htmlMetrics = metrics.map(m => `
                        <div class="bg-white p-5 rounded-xl shadow-md border-l-4 border-purple-600">
                            <div class="flex items-center">
                                ${m.icon}
                                <div class="ml-4">
                                    <h3 class="text-sm font-medium text-gray-500">${m.title}</h3>
                                    <p class="text-2xl font-bold text-gray-900">${m.value}</p>
                                </div>
                            </div>
                        </div>
                    `).join('');
                    document.getElementById('relatoriosMetrics').innerHTML = htmlMetrics;

                    // Tabela de Performance
                    let htmlPerf = data.performance.map(p => `
                        <tr class="hover:bg-gray-50">
                            <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${p.status || '-'}</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${p.status_cotacao || '-'}</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${p.quantidade}</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm font-bold text-gray-900">${formatCurrency(p.total)}</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${formatCurrency(p.media)}</td>
                        </tr>
                    `).join('');
                    document.getElementById('performanceBody').innerHTML = htmlPerf || `
                        <tr>
                            <td colspan="5" class="px-6 py-4 text-center text-gray-500">Nenhum dado de performance disponível.</td>
                        </tr>
                    `;
                });
        }
        
        function loadConfig() {
            fetch('/api/configuracoes')
                .then(r => r.json())
                .then(data => {
                    // Métricas de Configuração
                    const metrics = [
                        { title: 'Total de Registros', value: data.total_registros, icon: '<svg class="w-6 h-6 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 7v10a2 2 0 002 2h12a2 2 0 002-2V9a2 2 0 00-2-2h-3m-2 4l-3 3m0 0l-3-3m3 3V4"></path></svg>' },
                        { title: 'Tamanho do Banco', value: `${data.tamanho_mb} MB`, icon: '<svg class="w-6 h-6 text-red-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 7v10a2 2 0 002 2h12a2 2 0 002-2V9a2 2 0 00-2-2h-3m-2 4l-3 3m0 0l-3-3m3 3V4"></path></svg>' },
                        { title: 'Colunas no DB', value: data.colunas.split(',').length, icon: '<svg class="w-6 h-6 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 10h16M4 14h16M4 18h16"></path></svg>' }
                    ];
                    
                    let htmlMetrics = metrics.map(m => `
                        <div class="bg-white p-5 rounded-xl shadow-md border-l-4 border-indigo-600">
                            <div class="flex items-center">
                                ${m.icon}
                                <div class="ml-4">
                                    <h3 class="text-sm font-medium text-gray-500">${m.title}</h3>
                                    <p class="text-2xl font-bold text-gray-900">${m.value}</p>
                                </div>
                            </div>
                        </div>
                    `).join('');
                    document.getElementById('configMetrics').innerHTML = htmlMetrics;

                    // Colunas Esperadas
                    let htmlCols = data.colunas_esperadas.map((col, i) => `${i+1}. <code>${col}</code>`).join('<br>');
                    document.getElementById('colunasEsperadas').innerHTML = htmlCols;
                });
        }
        
        function limparDados() {
            if (!confirm('Deseja realmente limpar TODOS os dados? Esta ação é irreversível!')) return;
            if (!confirm('CONFIRME NOVAMENTE: Todos os dados serão removidos!')) return;
            
            fetch('/api/limpar', { method: 'POST' })
                .then(r => r.json())
                .then(data => {
                    alert(data.message);
                    if (data.success) showPage('dashboard', document.querySelector('.nav-btn'));
                });
        }
        
        window.onload = function() {
            // Inicializa o dashboard e o botão ativo
            const initialButton = document.querySelector('.nav-btn.active');
            if (initialButton) {
                initialButton.classList.add('bg-blue-700', 'bg-blue-800');
            }
            loadDashboard();
        };
    </script>
</body>
</html>
"""

# --------------------
# Banco e Schema
# --------------------
DB_FILE = 'ordens_servico_completo.db'

def init_database():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ordens_servico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descricao_operacao TEXT, numero_oportunidade TEXT, numero_vta TEXT,
            numero_cotacao TEXT, numero_circuito TEXT, status_cotacao TEXT,
            denominacao_produto TEXT, quantidade INTEGER, status TEXT,
            valor_pedido_bruto REAL, criado_em DATE, emissor_ordem TEXT,
            nome_emissor_ordem TEXT, nome_gerente_contas TEXT, organizacao_vendas TEXT,
            canal_distribuicao TEXT, setor_atividade TEXT, item_sd TEXT,
            id_produto TEXT, tempo_contrato TEXT,
            data_importacao DATETIME DEFAULT CURRENT_TIMESTAMP,
            data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_database()

def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

# --------------------
# Funções utilitárias para mapping e datas
# --------------------
def mapear_status(status_excel):
    if pd.isna(status_excel) or status_excel == '':
        return 'Sem Status'
    status = str(status_excel).strip()
    status_map = {
        'concluído': 'Concluído', 'concluido': 'Concluído', 'finalizado': 'Concluído',
        'completo': 'Concluído', 'pendente': 'Pendente', 'aberto': 'Aberto',
        'liberado': 'Liberado', 'liberada': 'Liberado', 'aprovado': 'Aprovado',
        'em andamento': 'Em Andamento', 'processando': 'Em Andamento',
        'cancelado': 'Cancelado', 'rejeitado': 'Rejeitado'
    }
    return status_map.get(status.lower(), status)

def converter_data(data_excel):
    if pd.isna(data_excel) or data_excel == '':
        return None
    try:
        if isinstance(data_excel, str):
            # Tenta converter string para datetime
            for fmt in ['%d.%m.%Y %H:%M:%S', '%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%Y-%m-%d', '%Y-%m-%d %H:%M:%S']:
                try:
                    return datetime.strptime(data_excel, fmt).date()
                except ValueError:
                    continue
            # Tenta parsear com pandas se os formatos acima falharem
            try:
                return pd.to_datetime(data_excel, errors='coerce').date()
            except:
                return None
        elif isinstance(data_excel, (int, float)):
            # Excel serial date
            return (pd.to_datetime('1899-12-30') + pd.Timedelta(days=int(data_excel))).date()
        
        # Tenta converter qualquer outro tipo para data
        return pd.to_datetime(data_excel).date()
    except:
        return None

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --------------------
# Funções de Consulta e Exportação (Reutilizáveis)
# --------------------
def build_query_and_params(request_args, limit=None):
    busca = request_args.get('busca', '').strip()
    status_filter = request_args.get('status', '').strip()
    cotacao_filter = request_args.get('status_cotacao', '').strip()
    
    query = "SELECT * FROM ordens_servico WHERE 1=1"
    params = []
    
    if busca:
        query += " AND (nome_emissor_ordem LIKE ? OR descricao_operacao LIKE ?)"
        param = f'%{busca}%'
        params.extend([param, param])
        
    if status_filter:
        query += " AND status = ?"
        params.append(status_filter)
        
    if cotacao_filter:
        query += " AND status_cotacao = ?"
        params.append(cotacao_filter)
        
    query += " ORDER BY id DESC"
    
    if limit:
        query += f" LIMIT {limit}"
        
    return query, params

# --------------------
# Rotas
# --------------------
@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/dashboard')
def dashboard_data():
    conn = get_db_connection()
    
    # Métricas
    total = conn.execute('SELECT COUNT(*) as total FROM ordens_servico').fetchone()['total']
    concluidas = conn.execute("SELECT COUNT(*) as total FROM ordens_servico WHERE status = 'Concluído'").fetchone()['total']
    pendentes = conn.execute("SELECT COUNT(*) as total FROM ordens_servico WHERE status = 'Pendente' OR status IS NULL OR status = 'Aberto' OR status = 'Em Andamento'").fetchone()['total']
    valor_total = conn.execute('SELECT SUM(valor_pedido_bruto) as total FROM ordens_servico').fetchone()['total'] or 0
    ultima_atualizacao = conn.execute('SELECT MAX(data_importacao) as data FROM ordens_servico').fetchone()['data'] or 'N/A'
    
    # Gráfico de Status
    status_data = conn.execute('SELECT IFNULL(status,"Sem Status") as status, COUNT(*) as count FROM ordens_servico GROUP BY status ORDER BY count DESC').fetchall()
    status_labels = [row['status'] for row in status_data]
    status_values = [row['count'] for row in status_data]
    
    # Gráfico de Cotação
    cotacao_data = conn.execute('''
        SELECT IFNULL(status_cotacao,"Sem Cotação") as status_cotacao, SUM(IFNULL(valor_pedido_bruto,0)) as total 
        FROM ordens_servico 
        GROUP BY status_cotacao 
        ORDER BY total DESC LIMIT 10
    ''').fetchall()
    cotacao_labels = [row['status_cotacao'] for row in cotacao_data]
    cotacao_values = [row['total'] or 0 for row in cotacao_data]
    
    # Timeline
    timeline_data = conn.execute('''
        SELECT strftime('%Y-%m', criado_em) as mes_ano, COUNT(*) as count 
        FROM ordens_servico 
        WHERE criado_em IS NOT NULL 
        GROUP BY mes_ano ORDER BY mes_ano
    ''').fetchall()
    timeline_labels = [row['mes_ano'] for row in timeline_data]
    timeline_values = [row['count'] for row in timeline_data]
    
    # Top Clientes
    top_clientes = conn.execute('''
        SELECT IFNULL(nome_emissor_ordem,"Sem Nome") as nome_emissor_ordem, COUNT(*) as count 
        FROM ordens_servico 
        GROUP BY nome_emissor_ordem 
        ORDER BY count DESC LIMIT 10
    ''').fetchall()
    
    # Top Produtos
    top_produtos = conn.execute('''
        SELECT IFNULL(denominacao_produto,"Sem Produto") as denominacao_produto, COUNT(*) as count 
        FROM ordens_servico 
        GROUP BY denominacao_produto 
        ORDER BY count DESC LIMIT 10
    ''').fetchall()
    
    conn.close()
    
    return jsonify({
        'metricas': {
            'total': total,
            'concluidas': concluidas,
            'pendentes': pendentes,
            'valor_total': 'R$ {:,.2f}'.format(valor_total).replace(',', 'X').replace('.', ',').replace('X', '.'), # Formato BR
            'ultima_atualizacao': str(ultima_atualizacao)[:10] if ultima_atualizacao != 'N/A' else 'N/A'
        },
        'status_chart': {'labels': status_labels, 'values': status_values},
        'cotacao_chart': {'labels': cotacao_labels, 'values': cotacao_values},
        'timeline_chart': {'labels': timeline_labels, 'values': timeline_values},
        'top_clientes': [{'nome': row['nome_emissor_ordem'], 'count': row['count']} for row in top_clientes],
        'top_produtos': [{'nome': row['denominacao_produto'], 'count': row['count']} for row in top_produtos]
    })

@app.route('/api/filtros')
def get_filtros():
    conn = get_db_connection()
    
    # Obter todos os status únicos
    status_data = conn.execute('SELECT DISTINCT status FROM ordens_servico WHERE status IS NOT NULL AND status != "" ORDER BY status').fetchall()
    status_cotacao_data = conn.execute('SELECT DISTINCT status_cotacao FROM ordens_servico WHERE status_cotacao IS NOT NULL AND status_cotacao != "" ORDER BY status_cotacao').fetchall()
    
    conn.close()
    
    status_list = [row['status'] for row in status_data]
    status_cotacao_list = [row['status_cotacao'] for row in status_cotacao_data]
    
    return jsonify({
        'status': status_list,
        'status_cotacao': status_cotacao_list
    })

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'Nenhum arquivo enviado'})
    file = request.files['file']
    atualizar = request.form.get('atualizar', 'true') == 'true'
    
    if file.filename == '':
        return jsonify({'success': False, 'message': 'Arquivo sem nome'})
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'message': 'Tipo de arquivo não suportado. Envie .xls ou .xlsx'})
    
    filename = secure_filename(file.filename)
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(save_path)
    
    try:
        # Lê com pandas
        df = pd.read_excel(save_path, engine='openpyxl' if filename.lower().endswith('x') else None)
        # Normalizar nomes das colunas para uppercase sem espaços extremos
        df.columns = [str(c).strip() for c in df.columns]
        
        # Mapeamento de colunas aprimorado
        col_map = {}
        cols_upper = {str(c).upper(): c for c in df.columns}
        
        # Mapeamento de colunas do DB para as colunas do Excel
        MAPPING_KEYS = {
            'descricao_operacao': ['DESCR', 'DESCRI', 'OPERACAO'],
            'status_cotacao': ['STATUS COT', 'STATUS_COT', 'STATUSCOTACAO', 'COTACAO STATUS'],
            'status': ['STATUS'],
            'denominacao_produto': ['PRODUTO', 'DENOMINACAO PRODUTO'],
            'nome_emissor_ordem': ['CLIENTE', 'NOME', 'EMISSOR', 'NOME EMISSOR'],
            'valor_pedido_bruto': ['VALOR', 'TOTAL', 'PRECO', 'VALOR PEDIDO'],
            'criado_em': ['CRIADO', 'DATA', 'DATA CRIACAO'],
            'numero_circuito': ['NUMERO CIRCUITO', 'CIRCUITO'],
            'numero_cotacao': ['NUMERO COTACAO', 'COTACAO']
        }
        
        for db_col, excel_keys in MAPPING_KEYS.items():
            for key in excel_keys:
                key_upper = key.upper()
                if key_upper in cols_upper:
                    col_map[db_col] = cols_upper[key_upper]
                    break
                # Heurística de substring
                for col_name_upper, col_name_raw in cols_upper.items():
                    if key_upper in col_name_upper:
                        col_map[db_col] = col_name_raw
                        break
                if db_col in col_map:
                    break

        # Substitua por valores None onde não existam
        def safe_get(row, key):
            return row[col_map[key]] if (key in col_map and col_map[key] in row.index) else None

        conn = get_db_connection()
        cursor = conn.cursor()

        if atualizar:
            cursor.execute("DELETE FROM ordens_servico")
            conn.commit()

        inserted = 0
        for _, row in df.iterrows():
            # Leitura segura:
            descricao = safe_get(row, 'descricao_operacao')
            status_raw = safe_get(row, 'status')
            status = mapear_status(status_raw)
            status_cot = safe_get(row, 'status_cotacao')
            produto = safe_get(row, 'denominacao_produto')
            cliente = safe_get(row, 'nome_emissor_ordem')
            valor = safe_get(row, 'valor_pedido_bruto')
            
            # Normalizar valor
            valor_num = None
            if valor is not None and not pd.isna(valor):
                try:
                    valor_num = float(valor)
                except:
                    # Tenta remover caracteres não numéricos
                    try:
                        valor_str = str(valor).replace('R$', '').replace('.', '').replace(',', '.').strip()
                        valor_num = float(valor_str)
                    except:
                        valor_num = None

            criado_em_raw = safe_get(row, 'criado_em')
            criado_em_dt = None
            if criado_em_raw is not None and not pd.isna(criado_em_raw):
                d = converter_data(criado_em_raw)
                if d:
                    criado_em_dt = d.isoformat()
            
            numero_circuito = safe_get(row, 'numero_circuito')
            numero_cotacao = safe_get(row, 'numero_cotacao')

            cursor.execute('''
                INSERT INTO ordens_servico (
                    descricao_operacao, numero_oportunidade, numero_vta,
                    numero_cotacao, numero_circuito, status_cotacao,
                    denominacao_produto, quantidade, status,
                    valor_pedido_bruto, criado_em, emissor_ordem,
                    nome_emissor_ordem, nome_gerente_contas, organizacao_vendas,
                    canal_distribuicao, setor_atividade, item_sd,
                    id_produto, tempo_contrato
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ''', (
                descricao, None, None,
                numero_cotacao, numero_circuito, status_cot,
                produto, None, status,
                valor_num, criado_em_dt, None,
                cliente, None, None,
                None, None, None,
                None, None
            ))
            inserted += 1

        conn.commit()
        conn.close()

        return jsonify({'success': True, 'message': f'Arquivo processado com sucesso. {inserted} linhas inseridas.'})
    except Exception as e:
        # Tenta remover o arquivo se o processamento falhar
        if os.path.exists(save_path):
            os.remove(save_path)
        print(f"Erro detalhado no upload: {e}")
        return jsonify({'success': False, 'message': f'Erro ao processar arquivo. Verifique se a planilha está no formato correto. Detalhe: {e}'})

@app.route('/api/consultar')
def consultar():
    query, params = build_query_and_params(request.args, limit=1000)
    
    conn = get_db_connection()
    rows = conn.execute(query, params).fetchall()
    
    results = []
    for r in rows:
        row_dict = {k: r[k] for k in r.keys()}
        results.append(row_dict)
        
    conn.close()
    return jsonify({'total': len(results), 'resultados': results})

@app.route('/api/exportar')
def exportar():
    # Não aplica limite para exportação
    query, params = build_query_and_params(request.args, limit=None)
    
    conn = get_db_connection()
    
    # CORREÇÃO: Passar a query e os params separadamente para pd.read_sql_query
    # Isso resolve o erro de "Número incorreto de associações fornecidas"
    try:
        df = pd.read_sql_query(query, conn, params=params)
    except Exception as e:
        conn.close()
        # Se houver erro na consulta (ex: sintaxe SQL inválida), retorna um erro amigável
        return jsonify({'success': False, 'message': f'Erro ao executar consulta para exportação: {e}'}), 500
        
    conn.close()
    
    if df.empty:
        # Retorna um arquivo Excel com uma mensagem de que não há dados
        df = pd.DataFrame({'Mensagem': ['Nenhum registro encontrado com os filtros aplicados.']})
        
    # Formata as colunas de data e valor para melhor visualização no Excel
    if 'criado_em' in df.columns:
        df['criado_em'] = pd.to_datetime(df['criado_em'], errors='coerce').dt.strftime('%d/%m/%Y')
    if 'data_importacao' in df.columns:
        df['data_importacao'] = pd.to_datetime(df['data_importacao'], errors='coerce').dt.strftime('%d/%m/%Y %H:%M:%S')
    if 'valor_pedido_bruto' in df.columns:
        # Renomeia a coluna para indicar que é valor em R$
        df = df.rename(columns={'valor_pedido_bruto': 'valor_pedido_bruto (R$)'})
        
    # Cria um buffer em memória para o arquivo Excel
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Dados Filtrados', index=False)
    # Fecha o writer para salvar o conteúdo no buffer
    writer.close()
    output.seek(0)
    
    # Retorna o arquivo para download
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        download_name=f'consulta_os_filtrada_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
        as_attachment=True
    )

@app.route('/api/relatorios')
def relatorios():
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM ordens_servico", conn)
    conn.close()

    total = len(df)
    valor_total = df['valor_pedido_bruto'].sum() if 'valor_pedido_bruto' in df.columns and not df.empty else 0
    ticket_medio = df['valor_pedido_bruto'].mean() if 'valor_pedido_bruto' in df.columns and total > 0 else 0

    performance = []
    if 'status' in df.columns and 'status_cotacao' in df.columns and not df.empty:
        # Preenche NaN com strings para o agrupamento funcionar corretamente
        df['status'] = df['status'].fillna('Sem Status')
        df['status_cotacao'] = df['status_cotacao'].fillna('Sem Cotação')
        
        grouped = df.groupby(['status', 'status_cotacao'])
        for (st, sc), g in grouped:
            total_g = g['valor_pedido_bruto'].sum() if 'valor_pedido_bruto' in g.columns else 0
            media_g = g['valor_pedido_bruto'].mean() if 'valor_pedido_bruto' in g.columns else 0
            
            performance.append({
                'status': st,
                'status_cotacao': sc,
                'quantidade': int(len(g)),
                'total': float(total_g) if pd.notna(total_g) else 0,
                'media': float(media_g) if pd.notna(media_g) else 0
            })

    return jsonify({
        'metricas': {
            'total': total,
            'valor_total': float(valor_total) if pd.notna(valor_total) else 0,
            'ticket_medio': float(ticket_medio) if pd.notna(ticket_medio) else 0
        },
        'performance': performance
    })

@app.route('/api/configuracoes')
def configuracoes():
    conn = get_db_connection()
    cur = conn.cursor()
    total_registros = cur.execute("SELECT COUNT(*) as c FROM ordens_servico").fetchone()['c']
    tamanho_mb = 0
    try:
        tamanho_mb = round(os.path.getsize(DB_FILE) / (1024*1024), 2)
    except:
        tamanho_mb = 0
        
    # Colunas esperadas (para instrução ao usuário)
    colunas_esperadas = [
        'descricao_operacao','numero_cotacao','numero_circuito','status_cotacao',
        'denominacao_produto','valor_pedido_bruto','criado_em','nome_emissor_ordem',
        'status' # Adicionado status para clareza
    ]
    
    # Colunas reais no DB
    df = pd.read_sql_query("SELECT * FROM ordens_servico LIMIT 1", conn)
    conn.close()
    colunas_reais = df.columns.tolist() if not df.empty else []
    
    return jsonify({
        'total_registros': total_registros,
        'tamanho_mb': tamanho_mb,
        'colunas_esperadas': colunas_esperadas,
        'colunas': ', '.join(colunas_reais) if colunas_reais else 'Sem registros'
    })

@app.route('/api/limpar', methods=['POST'])
def limpar():
    try:
        conn = get_db_connection()
        conn.execute("DELETE FROM ordens_servico")
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Todos os dados foram apagados!'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Erro ao limpar dados: {e}'})

# --------------------
# Run
# --------------------
if __name__ == '__main__':
    # Garante que o DB está inicializado antes de rodar o app
    init_database() 
    app.run(debug=True, host='0.0.0.0', port=5000)
