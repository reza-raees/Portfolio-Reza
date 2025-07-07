--Extraction of 4decimal values of the quantities in each formula and set the flag of 4cifre to the related item inggredients 

select g.item_code ,g.quantity, g.formula_id, i.UOM_CODE , g.line_id , i.C_FOUR_DECIMAL_INGRD , g.ATTRIBUTE30
from fsformulaingred g
JOIN fsitem i ON i.item_code = g.item_code AND i.UOM_CODE IN ('ea', 'pcs')
WHERE g.quantity <> ROUND(g.quantity, 3)

--Truncating the value of QUANTITY of the UOM(ea,pcs)
UPDATE g
SET 
    g.quantity = FLOOR(g.quantity * 1000) / 1000.0,
    g.attribute30 = '1'
FROM fsformulaingred g
JOIN fsitem i ON i.item_code = g.item_code AND i.UOM_CODE IN ('ea', 'pcs')
WHERE g.quantity <> ROUND(g.quantity, 3)


--Updating the item flag 4Cifre
UPDATE i
SET 
	i.C_FOUR_DECIMAL_INGRD = 1
FROM fsitem i
JOIN fsformulaingred g ON i.item_code = g.item_code and g.attribute30 = '1'

select distinct i.item_code 
from fsitem i 
JOIN fsformulaingred g ON i.item_code = g.item_code and g.attribute30 = '1'



--Rounding the QUANTITIES for not in UOM(ea and pcs)
UPDATE g
SET 
    g.quantity = ROUND(CAST(g.quantity AS numeric(18,6)), 3),
    g.attribute30 = '1'
FROM fsformulaingred g
JOIN fsitem i 
    ON i.item_code = g.item_code 
   AND i.UOM_CODE NOT IN ('ea', 'pcs')
   and i.item_code not like 'EVAP'
JOIN fsformula f 
    ON f.formula_id = g.formula_id 
   AND f.logical_delete <> 1
WHERE ABS(g.quantity - ROUND(g.quantity, 3)) > 0.00001;




select * from fsformulaingred g 
join fsitem i on i.item_code = g.item_code and i.item_code not like 'EVAP'
join fsformula f on f.formula_id = g.formula_id and f.logical_delete <> 1
WHERE ABS(g.quantity - ROUND(g.quantity, 3)) > 0.00001 and g.UOM_CODE NOT IN ('ea', 'pcs') 


--Update the Attribute30 for updating the 4Cifre flag
UPDATE g
SET 
    g.attribute30 = ''

FROM fsformulaingred g where g.attribute30 = 1

select *
from fsformulaingred g 
join fsitem i on i.item_code = g.item_code 
where g.attribute30 = 1

select * from fsitem where C_FOUR_DECIMAL_INGRD = 1