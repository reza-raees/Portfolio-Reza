-- Certification parameter master list
WITH CertParams AS (
    SELECT value AS param_code
    FROM (VALUES 
        ('ALLERGEN_EGG')
    ) AS v(value)
),

-- Step 1: Item-level parameters where component_ind = 8
IngredientParams AS (
    SELECT
        f.formula_code,
        f.version,
        f.formula_id,
        i.item_code,
		i.QUANTITY,
		b.ATTRIBUTE17,
        c.param_code,
        b.PVALUE,
        m.component_ind,
		g.level
    FROM FSFORMULA f
    JOIN FSFORMULAINGRED i ON f.formula_id = i.formula_id
    JOIN FSITEM m ON m.item_code = i.item_code
	Cross join CertParams c
    LEFT outer JOIN FSITEMTECHPARAM b 
        ON b.item_code = i.item_code 
        AND b.param_code = c.param_code
	JOIN ALLERGEN_EGG_RECALC_PRD G ON (G.FORMULA_CODE = F.FORMULA_CODE and g.version = f.version)
										
	 
    WHERE m.component_ind = 8
	AND G.LEVEL = 7
	and i.LINE_TYPE <> 2
),

-- Step 2: Formula-level parameters if component_ind = 1
FormulaParams AS (
    SELECT
        f.formula_code,
        f.version,
        f.formula_id,
        i.item_code,
		i.QUANTITY,
		p.ATTRIBUTE17,
        c.param_code,
        p.PVALUE,
        m.component_ind,
		g.level
    FROM FSFORMULA f
    JOIN FSFORMULAINGRED i ON f.formula_id = i.formula_id
    JOIN FSITEM m ON m.item_code = i.item_code
	Cross join CertParams c
    left outer JOIN FSFORMULATECHPARAM p 
        ON p.formula_id = case when i.ITEM_FORMULA_ID > 0 then i.ITEM_FORMULA_ID else m.formula_id end
        AND p.param_code = c.param_code
	JOIN ALLERGEN_EGG_RECALC_PRD G ON (G.FORMULA_CODE = F.FORMULA_CODE and g.version = f.version)
			
    WHERE m.component_ind = 1
	AND G.LEVEL = 7
	and i.LINE_TYPE <> 2
),

-- Combine all values from both sources
AllParams AS (
    SELECT * FROM IngredientParams
    UNION ALL
    SELECT * FROM FormulaParams
)
-- Final aggregation
SELECT 
    m.formula_code,
	--m.item_code,
    m.version,		
    m.param_code,
	--m.ATTRIBUTE17,
	--m.QUANTITY,
	--m.level,
    CASE 
        WHEN SUM(CASE WHEN m.PVALUE IS NULL OR LEN(TRIM(m.PVALUE)) = 0 THEN 1 ELSE 0 END) > 0 THEN NULL
        WHEN SUM(CASE WHEN m.PVALUE = '6' THEN 1 ELSE 0 END) > 0 
             AND SUM(CASE WHEN m.PVALUE = '7' THEN 1 ELSE 0 END) > 0 THEN '7'
        WHEN SUM(CASE WHEN m.PVALUE = '6' THEN 1 ELSE 0 END) > 0 THEN '6'
        WHEN SUM(CASE WHEN m.PVALUE = '7' THEN 1 ELSE 0 END) > 0 THEN '7'
        ELSE '-1'
    END AS final_value,
    STRING_AGG(
        CASE 
            WHEN m.PVALUE IS NULL OR LEN(TRIM(m.PVALUE)) = 0 THEN 'NULL' 
            ELSE CAST(m.PVALUE AS VARCHAR) 
        END, ', '
    ) AS all_param_values,
    ROUND(
    SUM(
        CASE 
            WHEN m.PVALUE IN ('6', '7') 
                 AND m.ATTRIBUTE17 IS NOT NULL 
                 AND m.Quantity IS NOT NULL THEN 
                 CAST(m.Quantity AS FLOAT) * CAST(m.ATTRIBUTE17 AS FLOAT) / 1000
            ELSE Null
        END
    ), 
    8  
) AS allergen_ppm_max

FROM AllParams m
WHERE m.item_code not like '%EVAP%' --and m.formula_code + '\' + m.VERSION  LIKE '%WS023020201-NOV-M00000001\002%' 
GROUP BY	
    m.formula_code,
    m.version, 
	--m.item_code,
	--m.ATTRIBUTE17,
	--m.QUANTITY,
    m.param_code
	--m.level,
	
ORDER BY 
    m.formula_code, 
    m.param_code;


	/*select * from fsformula where formula_code like 'WS023020201-NOV-M00000001' and version like '002'
	select * from fsformulaingred where formula_id like '62613'
	select * from fsformulatechparam where formula_id like  '62613' and param_code like 'allergen_egg'
	select * from fsitem where item_code like '500230205'*/
	