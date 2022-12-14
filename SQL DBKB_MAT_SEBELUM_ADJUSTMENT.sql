USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'DBKB_MAT_SEBELUM_ADJUSTMENT')
BEGIN
    DROP PROC DBKB_MAT_SEBELUM_ADJUSTMENT
END
GO
create PROCEDURE DBKB_MAT_SEBELUM_ADJUSTMENT
@VLC_MASUKAN_TAHUN NVARCHAR(4)
AS
BEGIN
	DECLARE C_PEK_KEG_X2 CURSOR FOR SELECT DISTINCT KD_PEKERJAAN, KD_KEGIATAN FROM HRG_SATUAN
	WHERE KD_PROPINSI    = '12' AND
	      KD_DATI2       = '12' AND
		  THN_HRG_SATUAN = @VLC_MASUKAN_TAHUN  AND
		  KD_PEKERJAAN IN('21','22','23','24')
	DECLARE @VLC_KD_PEK  NVARCHAR(2),
	@VLC_KD_KEG          NVARCHAR(2),
	@VLN_NILAI_HRG_SAT    FLOAT,
	@VLN_JUMLAH           BIGINT

	OPEN C_PEK_KEG_X2
	FETCH C_PEK_KEG_X2 INTO  @VLC_KD_PEK,@VLC_KD_KEG
	WHILE @@FETCH_STATUS=0
	BEGIN
		SELECT @VLN_NILAI_HRG_SAT=SUM(HRG_SATUAN) FROM   HRG_SATUAN
		WHERE  KD_PROPINSI    = '12' AND
	   		   KD_DATI2       = '12' AND
	   		   THN_HRG_SATUAN = @VLC_MASUKAN_TAHUN AND
	   		   KD_PEKERJAAN   = @VLC_KD_PEK        AND
	   		   KD_KEGIATAN    = @VLC_KD_KEG

        --SELECT @VLN_JUMLAH=COUNT(*) NILAI_DBKB_MATERIAL FROM DBKB_MATERIAL
        SELECT @VLN_JUMLAH=COUNT(*) FROM DBKB_MATERIAL
		WHERE  KD_PROPINSI       = '12' AND
	   		   KD_DATI2          = '12' AND
	   		   THN_DBKB_MATERIAL = @VLC_MASUKAN_TAHUN AND
	   		   KD_PEKERJAAN      = @VLC_KD_PEK        AND
	   		   KD_KEGIATAN       = @VLC_KD_KEG

		IF @VLN_JUMLAH = 0 
		BEGIN
   		   INSERT INTO DBKB_MATERIAL VALUES('12','12',@VLC_MASUKAN_TAHUN,@VLC_KD_PEK,@VLC_KD_KEG,@VLN_NILAI_HRG_SAT)
   		END
		ELSE
		BEGIN
			UPDATE DBKB_MATERIAL SET    NILAI_DBKB_MATERIAL =  @VLN_NILAI_HRG_SAT
			WHERE  KD_PROPINSI       = '12' AND
	   		       KD_DATI2          = '12' AND
	   		       THN_DBKB_MATERIAL = @VLC_MASUKAN_TAHUN AND
	   		       KD_PEKERJAAN      = @VLC_KD_PEK        AND
	   		       KD_KEGIATAN       = @VLC_KD_KEG
	    END
	    FETCH NEXT FROM C_PEK_KEG_X2 INTO  @VLC_KD_PEK,@VLC_KD_KEG
	END
	CLOSE C_PEK_KEG_X2
	DEALLOCATE C_PEK_KEG_X2
IF @@ERROR<>0
BEGIN
	ROLLBACK TRANSACTION
END
ELSE
BEGIN
	COMMIT TRANSACTION
END  
END
