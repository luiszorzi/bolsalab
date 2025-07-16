% Lê o arquivo CSV
dados = readtable('resistencia_1min.csv');

% Renomeia as colunas para nomes mais apropriados
dados.Properties.VariableNames = {'DataHora', 'Resistencia'};

% Converte a coluna DataHora para o formato datetime
dados.DataHora = datetime(dados.DataHora, 'InputFormat', 'yyyy-MM-dd HH:mm:ss');

% Exibe as primeiras linhas dos dados
disp(dados);

% Plotando os dados com mais zoom (eixo Y menos amplo)
plot(dados.DataHora, dados.Resistencia, '-o', 'MarkerSize', 4);
xlabel('Data e Hora');
ylabel('Resistência (Ohms)');
title('Resistência medida pelo Fluke 8846A');
grid on;

% Ajusta os limites do eixo Y para uma escala mais estreita
ylim([min(dados.Resistencia)-0.5, max(dados.Resistencia)+0.5]);

% Formatação do eixo X (opcional)
xtickformat('HH:mm:ss');
