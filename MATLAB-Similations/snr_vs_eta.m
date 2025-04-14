%This report presents a comprehensive analysis of the signal-to-noise ratio (SNR) in a dispersion-managed optical link, focusing on the influence of key physical parameters 
% such as transmitted power and span count. The study examines the degradation of SNR due to fiber impairments, including nonlinear effects, dispersion, and amplified spontaneous emission (ASE) noise. 
% By varying the optical power and the number of spans in the system, we calculate and model the relationship between SNR and the product of power and span count (η = P * N). 
% The findings are supported by two sets of simulations: SNR versus power for different span numbers, and SNR versus η. The results demonstrate a clear non-linear behavior between η and SNR, 
% with the SNR increasing at low η values and decreasing beyond an optimal power-span product. These results offer insights into the optimal power levels and span configurations 
% to minimize fiber impairments such as nonlinearity, dispersion, and amplified spontaneous emission (ASE) noise.

%   *******   (SNR) vs (eta = Power . Nspan)  **********


clear;
clc;

%% Global parameters
Nsymb = 1024;           % number of symbols
Nt = 32;                % number of discrete points per symbol

%% Tx parameters
symbrate = 10;          % symbol rate [Gbaud].
tx.rolloff = 0.2;       % pulse roll-off
tx.emph = 'asin';       % digital-premphasis type
modfor = 'ook';         % modulation format
lam = 1550;             % carrier wavelength [nm]

%% Channel parameters
RDPS  = 0;                  % residual dispersion per span [ps/nm]

spanVec = [ 3 , 6 , 9 , 12 , 15 , 18 , 21 ,24 , 27 , 30 , 33 , 36 , 39 , 42 , 45 , 48 ];   % Number of spans to test
Pvec =    [0.30 , 0.55 , 0.8 , 1.05 , 1.30 , 1.55 , 1.8 , 2.05 , 2.30 , 2.55 , 2.8 , 3.05 , 3.3 , 3.55 , 3.8 , 4.05 ];    % Power in mW

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
inigstate(Nsamp, fs);       % initialize global variables: Nsamp and fs.

%% Loop for testing spans and powers
SNRdBhat_all = zeros(1, length(spanVec));  % Pre-allocate SNR array
etaVec = zeros(1, length(spanVec));        % Pre-allocate eta array

for sp = 1:length(spanVec)
    Nspan = spanVec(sp);
    Plin = Pvec(sp);         % Power in mW

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
    SNRdBhat_all(sp) = mean(SNRdBhat);

    % Calculate eta = Nspan * Power
    etaVec(sp) =10 *log10(Nspan * Plin);
end

%% Plot SNR vs eta
figure;
plot(etaVec, SNRdBhat_all, '-o', 'LineWidth', 1.5); % Make line thicker
xlabel('\eta (Nspan * Power [dBm])', 'FontWeight', 'bold'); % Bold x-axis label
ylabel('SNR [dB]', 'FontWeight', 'bold'); % Bold y-axis label
title('SNR vs \eta', 'FontWeight', 'bold'); % Bold title
set(gca, 'FontWeight', 'bold', 'LineWidth', 1.5); % Bold axis ticks and box
grid on;

