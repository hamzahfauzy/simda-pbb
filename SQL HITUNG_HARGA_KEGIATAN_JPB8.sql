USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'HITUNG_HARGA_KEGIATAN_JPB8')
BEGIN
    DROP PROC HITUNG_HARGA_KEGIATAN_JPB8
END
GO
create PROCEDURE HITUNG_HARGA_KEGIATAN_JPB8
@VLC_THN_HRG_PEKERJAAN_JPB8 NVARCHAR(4)
AS
BEGIN
                            
	DECLARE @VLC_KD_PEKERJAAN            NVARCHAR(2),
	@VLC_KD_KEGIATAN             NVARCHAR(2),
	@VLN_LBR_BENT_MIN_HRG_JPB8   INT,
	@VLN_LBR_BENT_MAX_HRG_JPB8   INT,
	@VLN_TING_KOLOM_MIN_HRG_JPB8 INT,
	@VLN_TING_KOLOM_MAX_HRG_JPB8 INT,
	@VLN_HRG_KEGIATAN_JPB8       FLOAT,--NUMERIC(12,4),
	@VLN_NILAI                   BIGINT
	DECLARE C_VOL_KEGIATAN_JPB8 CURSOR FOR
	SELECT DISTINCT KD_PEKERJAAN,KD_KEGIATAN,LBR_BENT_MIN_HRG_JPB8,LBR_BENT_MAX_HRG_JPB8,TING_KOLOM_MIN_HRG_JPB8,TING_KOLOM_MAX_HRG_JPB8
	FROM VOL_KEGIATAN_JPB8
	OPEN C_VOL_KEGIATAN_JPB8
	FETCH C_VOL_KEGIATAN_JPB8 INTO @VLC_KD_PEKERJAAN,@VLC_KD_KEGIATAN,
			 @VLN_LBR_BENT_MIN_HRG_JPB8,@VLN_LBR_BENT_MAX_HRG_JPB8,@VLN_TING_KOLOM_MIN_HRG_JPB8,@VLN_TING_KOLOM_MAX_HRG_JPB8;
	WHILE @@FETCH_STATUS=0
	BEGIN
		SELECT @VLN_HRG_KEGIATAN_JPB8=SUM(A.VOL_KEGIATAN_JPB8 * B.HRG_SATUAN)FROM VOL_KEGIATAN_JPB8 A,HRG_SATUAN B
		WHERE A.KD_PEKERJAAN            = B.KD_PEKERJAAN              AND
			  A.KD_KEGIATAN             = B.KD_KEGIATAN               AND
			  A.KD_PEKERJAAN            = @VLC_KD_PEKERJAAN            AND
			  A.KD_KEGIATAN             = @VLC_KD_KEGIATAN             AND
			  A.LBR_BENT_MIN_HRG_JPB8   = @VLN_LBR_BENT_MIN_HRG_JPB8   AND
			  A.LBR_BENT_MAX_HRG_JPB8   = @VLN_LBR_BENT_MAX_HRG_JPB8   AND
			  A.TING_KOLOM_MIN_HRG_JPB8 = @VLN_TING_KOLOM_MIN_HRG_JPB8 AND
			  A.TING_KOLOM_MAX_HRG_JPB8 = @VLN_TING_KOLOM_MAX_HRG_JPB8 AND
			  B.KD_PROPINSI             = '12' AND
			  B.KD_DATI2                = '12' AND
			  B.THN_HRG_SATUAN          = @VLC_THN_HRG_PEKERJAAN_JPB8

		SELECT @VLN_NILAI=COUNT(*)FROM HRG_KEGIATAN_JPB8
		WHERE KD_PROPINSI             = '12' AND
			  KD_DATI2                = '12' AND
			  THN_HRG_PEKERJAAN_JPB8  = @VLC_THN_HRG_PEKERJAAN_JPB8 AND
			  KD_PEKERJAAN            = @VLC_KD_PEKERJAAN           AND
			  KD_KEGIATAN             = @VLC_KD_KEGIATAN            AND
			  LBR_BENT_MIN_HRG_JPB8   = @VLN_LBR_BENT_MIN_HRG_JPB8  AND
			  LBR_BENT_MAX_HRG_JPB8   = @VLN_LBR_BENT_MAX_HRG_JPB8  AND
			  TING_KOLOM_MIN_HRG_JPB8 = @VLN_TING_KOLOM_MIN_HRG_JPB8

		IF @VLN_NILAI = 0 
		BEGIN
		   INSERT INTO HRG_KEGIATAN_JPB8(KD_PROPINSI,KD_DATI2,THN_HRG_PEKERJAAN_JPB8,KD_PEKERJAAN,KD_KEGIATAN,
					   LBR_BENT_MIN_HRG_JPB8,LBR_BENT_MAX_HRG_JPB8,TING_KOLOM_MIN_HRG_JPB8,TING_KOLOM_MAX_HRG_JPB8,HRG_KEGIATAN_JPB8)
		   VALUES ('12','12',@VLC_THN_HRG_PEKERJAAN_JPB8,@VLC_KD_PEKERJAAN,@VLC_KD_KEGIATAN,@VLN_LBR_BENT_MIN_HRG_JPB8,
				   @VLN_LBR_BENT_MAX_HRG_JPB8,@VLN_TING_KOLOM_MIN_HRG_JPB8,@VLN_TING_KOLOM_MAX_HRG_JPB8,@VLN_HRG_KEGIATAN_JPB8)
		END
		ELSE
		BEGIN
		   UPDATE HRG_KEGIATAN_JPB8 SET    HRG_KEGIATAN_JPB8=@VLN_HRG_KEGIATAN_JPB8
		   WHERE KD_PROPINSI             = '12' AND
				 KD_DATI2                = '12' AND
				 THN_HRG_PEKERJAAN_JPB8  = @VLC_THN_HRG_PEKERJAAN_JPB8  AND
				 KD_PEKERJAAN            = @VLC_KD_PEKERJAAN            AND
				 KD_KEGIATAN             = @VLC_KD_KEGIATAN             AND
				 LBR_BENT_MIN_HRG_JPB8   = @VLN_LBR_BENT_MIN_HRG_JPB8   AND
				 LBR_BENT_MAX_HRG_JPB8   = @VLN_LBR_BENT_MAX_HRG_JPB8   AND
				 TING_KOLOM_MIN_HRG_JPB8 = @VLN_TING_KOLOM_MIN_HRG_JPB8 AND
				 TING_KOLOM_MAX_HRG_JPB8 = @VLN_TING_KOLOM_MAX_HRG_JPB8
		END 
		FETCH NEXT FROM C_VOL_KEGIATAN_JPB8 INTO @VLC_KD_PEKERJAAN,@VLC_KD_KEGIATAN,
		@VLN_LBR_BENT_MIN_HRG_JPB8,@VLN_LBR_BENT_MAX_HRG_JPB8,@VLN_TING_KOLOM_MIN_HRG_JPB8,@VLN_TING_KOLOM_MAX_HRG_JPB8;
	  END
	  CLOSE C_VOL_KEGIATAN_JPB8
	  DEALLOCATE C_VOL_KEGIATAN_JPB8
if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
End	  