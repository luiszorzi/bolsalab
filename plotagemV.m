% Lê o arquivo CSV de tensão
dados_v = readtable('tensao_1min.csv');

% Renomeia as colunas
dados_v.Properties.VariableNames = {'DataHora', 'Tensao'};

% Converte a coluna DataHora para datetime
dados_v.DataHora = datetime(dados_v.DataHora, 'InputFormat', 'yyyy-MM-dd HH:mm:ss');

% Exibe as primeiras linhas dos dados
disp(dados_v);

% Plotando os dados de tensão
plot(dados_v.DataHora, dados_v.Tensao, '-o', 'MarkerSize', 4);
xlabel('Data e Hora');
ylabel('Tensão (V)');
title('Tensão medida pelo Fluke 8846A');
grid on;

% Ajusta os limites do eixo Y com margem
ylim([min(dados_v.Tensao)-0.5, max(dados_v.Tensao)+0.5]);

% Formatação do eixo X
xtickformat('HH:mm:ss');
