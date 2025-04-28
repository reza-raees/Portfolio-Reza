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
	AND G.LEVEL = 9
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
	AND G.LEVEL = 9
),

-- Combine all values from both sources
AllParams AS (
    SELECT * FROM IngredientParams
    UNION ALL
    SELECT * FROM FormulaParams
)

, FinalAggregation AS (
	SELECT
		--formula_code,
		--Version,
		formula_id,
		param_code,
		CASE 
			WHEN SUM(CASE WHEN PVALUE IS NULL THEN 1 ELSE 0 END) > 0 THEN NULL
			ELSE MIN(CONVERT(INT, PVALUE))
		END AS min_param_value
	FROM AllParams 
	where item_code not like '%EVAP%' --and formula_code + '\' + VERSION  LIKE '%C0160023-NOV-M00000001\007%' 
	GROUP BY 
		formula_id, 
		param_code
		--formula_code,
		--Version
)

MERGE FSFORMULATECHPARAM AS target
USING (
	SELECT
		formula_id,
		param_code,
		CAST(min_param_value AS VARCHAR) AS PVALUE,
		CASE 
			WHEN min_param_value IS NULL THEN 0 
			ELSE 2 
		END AS CALC_LEVEL
	FROM FinalAggregation
	
) AS source
ON target.formula_id = source.formula_id
   AND target.param_code = source.param_code

WHEN MATCHED THEN
	UPDATE SET 
		target.PVALUE = source.PVALUE,
		target.CALC_LEVEL = source.CALC_LEVEL

WHEN NOT MATCHED BY TARGET THEN
	INSERT (
		FORMULA_ID, PARAM_CODE, CALC_LEVEL, PVALUE,
		DECDIGIT, TEST_CODE, FS_SYS_ROWID, ROLLUP_IND,
		ATTRIBUTE1, ATTRIBUTE2, ATTRIBUTE3, ATTRIBUTE4, ATTRIBUTE5,
		ATTRIBUTE6, ATTRIBUTE7, ATTRIBUTE8, ATTRIBUTE9, ATTRIBUTE10,
		DISPLAY_UM, ATTRIBUTE11, ATTRIBUTE12, ATTRIBUTE13, ATTRIBUTE14, ATTRIBUTE15,
		ATTRIBUTE16, ATTRIBUTE17, ATTRIBUTE18, ATTRIBUTE19, ATTRIBUTE20, ATTRIBUTE21,
		ATTRIBUTE22, ATTRIBUTE23, ATTRIBUTE24, ATTRIBUTE25, MAXVAL, MINVAL
	)
	VALUES (
		source.FORMULA_ID, source.param_code, source.CALC_LEVEL , source.PVALUE,
		NULL, NULL, NULL, NULL,
		NULL, NULL, NULL, NULL, NULL,
		NULL, NULL, NULL, NULL, NULL,
		NULL, NULL, NULL, NULL, NULL, NULL,
		NULL, NULL, NULL, NULL, NULL, NULL,
		NULL, NULL, NULL, NULL, NULL, NULL
	);



	
