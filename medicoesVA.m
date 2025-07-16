% Nome do arquivo CSV (mesmo caminho onde salvou as medições)
filename = 'C:\Users\luisp\Downloads\projeto\dados_medicoes.csv';

% Lê os dados do CSV, ignorando a primeira linha (cabeçalho)
data = readmatrix(filename, 'NumHeaderLines', 1);

tensao = data(:,1);
corrente = data(:,2);

nPontos = length(tensao);
tempo = (0:nPontos-1) * 5;  % tempo aproximado, considerando intervalo de 5 segundos

% Plot tensão
figure;
subplot(2,1,1);
plot(tempo, tensao, '-o', 'LineWidth', 1.5);
grid on;
xlabel('Tempo (s)');
ylabel('Tensão (V)');
title('Tensão medida vs. Tempo');

% Plot corrente
subplot(2,1,2);
plot(tempo, corrente, '-o', 'LineWidth', 1.5);
grid on;
xlabel('Tempo (s)');
ylabel('Corrente (A)');
title('Corrente medida vs. Tempo');
