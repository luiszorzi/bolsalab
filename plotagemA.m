% Lê o arquivo CSV com os dados
dados = readtable('dados_medicoes.csv');

% Renomeia as colunas, se necessário
% Supondo que o CSV tem colunas: Tempo (s), Tensao (V), Corrente (A)
dados.Properties.VariableNames = {'Tempo', 'Tensao', 'Corrente'};

% Exibe as primeiras linhas da tabela
disp(dados);

% === Gráfico de Tensão ===
subplot(2,1,1);  % Primeiro subplot
plot(dados.Tempo, dados.Tensao, '-o', 'MarkerSize', 4);
xlabel('Tempo (s)');
ylabel('Tensão (V)');
title('Tensão medida pela fonte');
grid on;
ylim([min(dados.Tensao)-0.5, max(dados.Tensao)+0.5]);

% === Gráfico de Corrente ===
subplot(2,1,2);  % Segundo subplot
plot(dados.Tempo, dados.Corrente, '-o', 'MarkerSize', 4);
xlabel('Tempo (s)');
ylabel('Corrente (A)');
title('Corrente medida pela fonte');
grid on;
ylim([min(dados.Corrente)-0.01, max(dados.Corrente)+0.01]);
