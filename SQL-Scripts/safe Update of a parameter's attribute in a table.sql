SELECT
    p.item_code,
    p.param_code,
    p.attribute16 AS original_attribute16,
    STUFF(
        (
            SELECT ';' + 
                   ISNULL(v.enum_value, LTRIM(RTRIM(s2.value)))
            FROM dbo.split_string(p.attribute16, ',') AS s2
            LEFT JOIN FSVALIDENUMLABELCF l2
                ON LTRIM(RTRIM(s2.value)) = l2.enum_label
                AND l2.enum_code = 'C_COUNTRIES2'
                AND l2.language_code = 'IT-IT'
            LEFT JOIN FSVALIDENUMVALCF v
                ON l2.enum_code = v.enum_code
                AND l2.enum_order = v.enum_order
            FOR XML PATH(''), TYPE
        ).value('.', 'NVARCHAR(MAX)')
        , 1, 1, ''
    ) AS new_attribute16
FROM FSITEMTECHPARAM p
WHERE p.param_code = 'GEOGRAPHICAL_ORIGIN'
  AND p.attribute16 IS NOT NULL
  AND p.attribute16 <> ''
ORDER BY p.item_code;



UPDATE p
SET p.attribute16 =
    STUFF(
        (
            SELECT ';' + 
                   ISNULL(v.enum_value, LTRIM(RTRIM(s2.value)))
            FROM dbo.split_string(p.attribute16, ',') AS s2
            LEFT JOIN FSVALIDENUMLABELCF l2
                ON LTRIM(RTRIM(s2.value)) = l2.enum_label
                AND l2.enum_code = 'C_COUNTRIES2'
                AND l2.language_code = 'IT-IT'
            LEFT JOIN FSVALIDENUMVALCF v
                ON l2.enum_code = v.enum_code
                AND l2.enum_order = v.enum_order
            FOR XML PATH(''), TYPE
        ).value('.', 'NVARCHAR(MAX)')
        , 1, 1, ''
    )
FROM FSITEMTECHPARAM p
WHERE p.param_code = 'GEOGRAPHICAL_ORIGIN'
  AND p.attribute16 IS NOT NULL
  AND p.attribute16 <> '';