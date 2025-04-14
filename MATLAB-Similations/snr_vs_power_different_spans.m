
% ********   SNR vs Power (in different Nspan)   *********


clear
clc

%% Global parameters

Nsymb = 1024;           % number of symbols
Nt = 32;                % number of discrete points per symbol

%% Tx parameters
symbrate = 10;          % symbol rate [Gbaud].
tx.rolloff = 0.2;       % pulse roll-off
tx.emph = 'asin';       % digital-premphasis type
modfor = 'ook';         % modulation format
Pvec = -5:5:30;         % power [dBm] from -5 dBm to 30 dBm
%Pvec = 0;
lam = 1550;             % carrier wavelength [nm]

%% Channel parameters

RDPS  = 0;                  % residual dispersion per span [ps/nm]
spanVec = [1 , 5 , 10 , 25 , 60];   % different span numbers to test

% Transmission fiber
ft.length     = 100E3;      % length [m]
ft.lambda     = 1550;       % wavelength [nm] of fiber parameters
ft.alphadB    = 0.2;        % attenuation [dB/km]
ft.disp       = 17;         % dispersion [ps/nm/km] @ ft.lambda
ft.slope      = 0.057;      % slope [ps/nm^2/km] @ ft.lambda
ft.pmdpar     = 0;          % PMD parameter [ps/sqrt(km)]
ft.ismanakov  = true;       % Solve Manakov equation
ft.aeff       = 80;         % effective area [um^2]
ft.n2         = 2.5e-20;    % nonlinear index [m^2/W]
ft.dzmax      = 2E4;        % maximum SSFM step size [m]
ft.trace      = true;       % show information on screen

% compensating fiber
fc = ft;                    % same parameters but:
fc.length     = 1e3;        % fixed length [m]
fc.alphadB    = 0.6;        % attenuation [dB/km]
fc.disp       = (RDPS*1e3 - ft.disp*ft.length)/fc.length; % [ps/nm/km] to get RDPS
fc.slope      = 0.057;      % slope [ps/nm^2/km] @ fc.lambda
fc.pmdpar     = 0;          % PMD parameter [ps/sqrt(km)]
fc.aeff       = 20;         % effective area [um^2]
fc.dzmax      = 2E4;        % maximum step size [m]

% Optical amplifier
amp.f = 6;														   

%% Rx parameters
rx.modformat = modfor;      % modulation format
rx.sync.type = 'da';        % time-recovery method
rx.oftype = 'gauss';        % optical filter type
rx.obw = Inf;               % optical filter bandwidth normalized to symbrate
rx.eftype = 'rootrc';       % electrical filter type
rx.ebw = 0.5;               % electrical filter bandwidth normalized to symbrate
rx.epar = tx.rolloff;
rx.type = 'bin';            % binary pattern

%% Init
Nsamp = Nsymb*Nt;           % overall number of samples
fs = symbrate*Nt;           % sampling rate [GHz]
inigstate(Nsamp,fs);        % initialize global variables: Nsamp and fs.

%% Loop for testing different spans and power levels
SNRdBhat_all = zeros(length(spanVec), length(Pvec));  % Pre-allocate SNR array

for sp = 1:length(spanVec)
    Nspan = spanVec(sp);
    
    for p = 1:length(Pvec)
        PdBm = Pvec(p);     % Current power [dBm]
        Plin = 10.^(PdBm/10);  % Power [mW]

        % Tx side
        E = lasersource(Plin, lam);  							
        rng(1);
        [patx, patbinx] = datapattern(Nsymb, 'rand', struct('format', modfor));
        [sigx, normx] = digitalmod(patbinx, modfor, symbrate, 'costails', tx);
        E = iqmodulator(E, sigx, struct('norm', normx));

        % Channel for the current span
        SNRdBhat = zeros(1, Nspan);  % SNR for each span
        
        for ns = 1:Nspan
            [E, parft] = fiber(E, ft);
            [E, parfc] = fiber(E, fc);
            amp.gain = (ft.length*ft.alphadB + fc.length*fc.alphadB)*1e-3; % Adjust amp gain
            E = ampliflat(E, amp);
            
            % Rx side
            rsig = rxfrontend(E, lam, symbrate, rx);	
            akhat = rxdsp(rsig, symbrate, patx, rx);  
            patbinhat = samp2pat(akhat, modfor, rx); 
            
            % SNR calculation
            SNRdBhat(ns) = samp2snr(patbinx, akhat, modfor);  
        end

        % Average SNR over the spans
        SNRdBhat_all(sp, p) = mean(SNRdBhat);
    end
end

%% Plot SNR vs Power for different spans
figure;
hold on;
lineStyles = {'-', '--', ':', '-.', '-'};
markers = {'o', 's', 'd', '^', 'v'}; % Different markers for each line
for sp = 1:length(spanVec)
    
    plot(Pvec, SNRdBhat_all(sp, :), 'LineStyle', lineStyles{sp}, ...
         'Marker', markers{sp}, 'Color', 'k', 'DisplayName', ['Nspan = ', num2str(spanVec(sp))]);
end
hold off;
xlabel('Power [dBm]');
ylabel('SNR [dB]');
title('SNR vs Power for Different Span Numbers');
legend show;
grid on;
