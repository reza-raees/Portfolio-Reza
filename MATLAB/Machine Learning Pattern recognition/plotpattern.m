function plotpatternupdated(pat1, pat2)
    % Visualize two binary patterns as images

    % Create a colormap with grayscale colors
    map = zeros(256, 3);
    for i = 0:255
        map(i+1,:) = [i/255, i/255, i/255];
    end
    colormap(map);

    % Initialize buffers for the patterns
    buf1 = zeros(13, 8);
    buf2 = zeros(13, 8);

    % Populate the buffers with pattern data
    count = 1;
    for j = 1:13
        for k = 1:8
            if count <= numel(pat1)
                buf1(j, k) = pat1(count);
            end
            if count <= numel(pat2)
                buf2(j, k) = pat2(count);
            end
            count = count + 1;
        end
    end

    % Display the first pattern
    subplot(1, 2, 1);
    image(255 * (1 - buf1));
    title('Pattern 1');

    % Display the second pattern
    subplot(1, 2, 2);
    image(255 * (1 - buf2));
    title('Pattern 2');
end
