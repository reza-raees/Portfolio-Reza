-- Certification parameter master list
WITH CertParams AS (
    SELECT value AS param_code
    FROM (VALUES 
        ('HALAL'), ('HALAL_PACK_LOGO'), ('KOSHER'), ('KOSHER_PASSOVER'), 
        ('KOSHER_PACK_LOGO'), ('BIO'), ('FAIRTRADE'), ('RSPO_MASSBALANCE'), 
        ('VEGAN'), ('VEGETARIAN'), ('UTZ'), ('IP'), ('IGP'), ('DOP'), 
        ('VLOG'), ('AIC'), ('RSPO_SEGREGATED')
    ) AS v(value)
),

-- Step 1: Item-level parameters where component_ind = 8
IngredientParams AS (
    SELECT
        f.formula_code,
        f.version,
        f.formula_id,
        i.item_code,
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
	JOIN GLOBALCALC_SO_CERT_CALC_TST G ON (G.FORMULA_CODE = F.FORMULA_CODE and g.version = f.version)
										
	 
    WHERE m.component_ind = 8
	AND G.LEVEL = 2
),

-- Step 2: Formula-level parameters if component_ind = 1
FormulaParams AS (
    SELECT
        f.formula_code,
        f.version,
        f.formula_id,
        i.item_code,
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
	JOIN GLOBALCALC_SO_CERT_CALC_TST G ON (G.FORMULA_CODE = F.FORMULA_CODE and g.version = f.version)
			
    WHERE m.component_ind = 1
	AND G.LEVEL = 2
),

-- Combine all values from both sources
AllParams AS (
    SELECT * FROM IngredientParams
    UNION ALL
    SELECT * FROM FormulaParams
)
-- Final aggregation
SELECT  --TOP 50
    m.formula_code,
	--m.item_code,
    m.version,		
    m.param_code,
	--m.level,
    CASE 
        WHEN SUM(CASE WHEN m.PVALUE IS NULL or len(trim(m.pvalue)) = 0 THEN 1 ELSE 0 END) > 0 THEN NULL
        ELSE MIN(convert(int, m.PVALUE))
    END AS min_param_value,
    STRING_AGG(
		CASE WHEN m.PVALUE IS NULL or len(trim(m.pvalue)) = 0 THEN 'NULL' else CAST(m.PVALUE AS VARCHAR) end, ', '
        --ISNULL(CAST(a.PVALUE AS VARCHAR), 'NULL'), ', '
    ) AS all_param_values
FROM AllParams m
WHERE m.item_code not like '%EVAP%' --and m.formula_code + '\' + m.VERSION  LIKE '%C0160023-NOV-M00000001\007%' and m.param_code = 'halal'
GROUP BY	
    m.formula_code, 
    m.version, 
    m.param_code
	--m.level,
	--m.item_code			   
ORDER BY 
    m.formula_code, 
    m.param_code;




	