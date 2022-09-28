USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'H_KEGIATAN')
BEGIN
    DROP PROC H_KEGIATAN
END
GO

CREATE PROCEDURE H_KEGIATAN
@VLC_THN_HRG_SATUAN nvarchar(4)
AS
BEGIN
  DECLARE @VLC_KD_PEKERJAAN  NVARCHAR(2),
  @VLC_KD_KEGIATAN   NVARCHAR(2),
  @VLC_KD_JPB        NVARCHAR(2),
  @VLC_KD_BNG_LANTAI NVARCHAR(8),
  @VLC_TIPE_BNG      NVARCHAR(5),
  @VLN_HRG_KEGIATAN   FLOAT,
  @VLN_NILAI          INT
  DECLARE C_PEK_KEG CURSOR FOR
       SELECT DISTINCT KD_PEKERJAAN,KD_KEGIATAN,KD_JPB,KD_BNG_LANTAI,TIPE_BNG
       FROM VOL_KEGIATAN
  OPEN C_PEK_KEG
  FETCH C_PEK_KEG INTO @VLC_KD_PEKERJAAN,@VLC_KD_KEGIATAN,@VLC_KD_JPB,
    @VLC_KD_BNG_LANTAI,@VLC_TIPE_BNG
  WHILE @@FETCH_STATUS = 0
  BEGIN
    SELECT @VLN_HRG_KEGIATAN=ROUND(ROUND(A.VOL_KEGIATAN,2) * ROUND(B.HRG_SATUAN,2),2)--ROUND((A.VOL_KEGIATAN * B.HRG_SATUAN),2)
	FROM VOL_KEGIATAN A,HRG_SATUAN B
    WHERE A.KD_PEKERJAAN=B.KD_PEKERJAAN AND A.KD_KEGIATAN=B.KD_KEGIATAN AND
          A.KD_PEKERJAAN=@VLC_KD_PEKERJAAN AND A.KD_KEGIATAN=@VLC_KD_KEGIATAN AND
          A.KD_JPB=@VLC_KD_JPB AND A.KD_BNG_LANTAI=@VLC_KD_BNG_LANTAI AND A.TIPE_BNG=@VLC_TIPE_BNG AND
          B.KD_PROPINSI='12' AND B.KD_DATI2='12' AND
          B.THN_HRG_SATUAN=@VLC_THN_HRG_SATUAN;
	IF @VLN_HRG_KEGIATAN IS NULL 
	BEGIN
		SET @VLN_HRG_KEGIATAN = 0
	END
	
    SELECT @VLN_NILAI=COUNT(*) FROM HRG_KEGIATAN
    WHERE KD_PROPINSI='12' AND KD_DATI2='12' AND
          THN_KEGIATAN=@VLC_THN_HRG_SATUAN AND KD_JPB=@VLC_KD_JPB AND TIPE_BNG=@VLC_TIPE_BNG AND
          KD_BNG_LANTAI=@VLC_KD_BNG_LANTAI AND KD_PEKERJAAN=@VLC_KD_PEKERJAAN AND KD_KEGIATAN=@VLC_KD_KEGIATAN
    IF @VLN_NILAI = 0 
    BEGIN
       INSERT INTO HRG_KEGIATAN(KD_PROPINSI,KD_DATI2,THN_KEGIATAN,KD_JPB,TIPE_BNG,KD_BNG_LANTAI,KD_PEKERJAAN,KD_KEGIATAN,HRG_KEGIATAN)
       VALUES ('12','12',@VLC_THN_HRG_SATUAN,@VLC_KD_JPB,@VLC_TIPE_BNG,@VLC_KD_BNG_LANTAI,@VLC_KD_PEKERJAAN,@VLC_KD_KEGIATAN,@VLN_HRG_KEGIATAN)
     END
    ELSE
    BEGIN
       UPDATE HRG_KEGIATAN SET HRG_KEGIATAN=@VLN_HRG_KEGIATAN
       WHERE KD_PROPINSI='12' AND KD_DATI2='12' AND
       THN_KEGIATAN=@VLC_THN_HRG_SATUAN AND KD_JPB=@VLC_KD_JPB AND TIPE_BNG=@VLC_TIPE_BNG AND
       KD_BNG_LANTAI=@VLC_KD_BNG_LANTAI AND KD_PEKERJAAN=@VLC_KD_PEKERJAAN AND KD_KEGIATAN=@VLC_KD_KEGIATAN
    END
   FETCH NEXT FROM C_PEK_KEG INTO @VLC_KD_PEKERJAAN,@VLC_KD_KEGIATAN,@VLC_KD_JPB,@VLC_KD_BNG_LANTAI,@VLC_TIPE_BNG
  END 
  CLOSE C_PEK_KEG
  DEALLOCATE C_PEK_KEG
IF @@ERROR<>0
BEGIN
	ROLLBACK TRAN
END
ELSE
BEGIN
	COMMIT TRAN
END  
END 