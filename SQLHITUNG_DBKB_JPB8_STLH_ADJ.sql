USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'HITUNG_DBKB_JPB8_STLH_ADJ')
BEGIN
    DROP PROC HITUNG_DBKB_JPB8_STLH_ADJ
END
GO
create PROCEDURE HITUNG_DBKB_JPB8_STLH_ADJ
@VLC_THN_DBKB_JPB8  NVARCHAR(4)
AS
BEGIN                           
	DECLARE 
	@VLN_LBR_BENT_MIN_DBKB_JPB8    INT,
	@VLN_LBR_BENT_MAX_DBKB_JPB8    INT,
	@VLN_TING_KOLOM_MIN_DBKB_JPB8  INT,
	@VLN_TING_KOLOM_MAX_DBKB_JPB8  INT,
	@VLN_NILAI_DBKB_JPB8           FLOAT,
	@VLC_KD_ADJ                    NVARCHAR(2),
	@VLN_LBR_BENT_MIN_ADJ          INT,
	@VLN_LBR_BENT_MAX_ADJ          INT,
	@VLN_TING_KOLOM_MIN_ADJ        INT,
	@VLN_TING_KOLOM_MAX_ADJ        INT,
	@VLN_PCT_ADJ_BNG_JPB8          FLOAT,
	@VLN_DBKB_JPB8_STLH_ADJ        FLOAT--NUMERIC(12,2)
	DECLARE C_DBKB_JPB8 CURSOR FOR
	SELECT DISTINCT A.LBR_BENT_MIN_DBKB_JPB8, A.LBR_BENT_MAX_DBKB_JPB8, A.TING_KOLOM_MIN_DBKB_JPB8, A.TING_KOLOM_MAX_DBKB_JPB8, A.NILAI_DBKB_JPB8
	FROM DBKB_JPB8 A,HRG_KEGIATAN_JPB8 B
	WHERE A.KD_PROPINSI              = B.KD_PROPINSI             AND
				 A.KD_DATI2                 = B.KD_DATI2                AND
				 A.THN_DBKB_JPB8            = B.THN_HRG_PEKERJAAN_JPB8  AND
				 A.LBR_BENT_MIN_DBKB_JPB8   = B.LBR_BENT_MIN_HRG_JPB8   AND
				 A.LBR_BENT_MAX_DBKB_JPB8   = B.LBR_BENT_MAX_HRG_JPB8   AND
				 A.TING_KOLOM_MIN_DBKB_JPB8 = B.TING_KOLOM_MIN_HRG_JPB8 AND
				 A.TING_KOLOM_MAX_DBKB_JPB8 = B.TING_KOLOM_MAX_HRG_JPB8 AND
				 A.KD_PROPINSI              = '12' AND
				 A.KD_DATI2                 = '12' AND
				 A.THN_DBKB_JPB8            = @VLC_THN_DBKB_JPB8
	OPEN C_DBKB_JPB8
	FETCH C_DBKB_JPB8 INTO @VLN_LBR_BENT_MIN_DBKB_JPB8,@VLN_LBR_BENT_MAX_DBKB_JPB8,
			@VLN_TING_KOLOM_MIN_DBKB_JPB8,@VLN_TING_KOLOM_MAX_DBKB_JPB8,@VLN_NILAI_DBKB_JPB8
	WHILE @@FETCH_STATUS =0
	BEGIN
		  --NILAI DBKB JPB 8 STLH ADJUSTMEN
		  SET @VLN_DBKB_JPB8_STLH_ADJ = @VLN_NILAI_DBKB_JPB8 * (1 + 0.27)
		  SET @VLN_DBKB_JPB8_STLH_ADJ = FLOOR(@VLN_DBKB_JPB8_STLH_ADJ)

		  UPDATE DBKB_JPB8 SET NILAI_DBKB_JPB8 = @VLN_DBKB_JPB8_STLH_ADJ
		  WHERE KD_PROPINSI              = '12' AND
				KD_DATI2                 = '12' AND
				THN_DBKB_JPB8            = @VLC_THN_DBKB_JPB8            AND
				LBR_BENT_MIN_DBKB_JPB8   = @VLN_LBR_BENT_MIN_DBKB_JPB8   AND
				LBR_BENT_MAX_DBKB_JPB8   = @VLN_LBR_BENT_MAX_DBKB_JPB8   AND
				TING_KOLOM_MIN_DBKB_JPB8 = @VLN_TING_KOLOM_MIN_DBKB_JPB8 AND
				TING_KOLOM_MAX_DBKB_JPB8 = @VLN_TING_KOLOM_MAX_DBKB_JPB8
		FETCH NEXT FROM C_DBKB_JPB8 INTO @VLN_LBR_BENT_MIN_DBKB_JPB8,@VLN_LBR_BENT_MAX_DBKB_JPB8,
			@VLN_TING_KOLOM_MIN_DBKB_JPB8,@VLN_TING_KOLOM_MAX_DBKB_JPB8,@VLN_NILAI_DBKB_JPB8	  
	END
	   CLOSE C_DBKB_JPB8
	   DEALLOCATE C_DBKB_JPB8
if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
END
	   
	   