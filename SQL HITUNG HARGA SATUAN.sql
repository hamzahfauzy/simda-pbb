USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'H_SATUAN')
BEGIN
    DROP PROC H_SATUAN
END
GO
CREATE proc H_SATUAN
@xTahun nvarchar(4)
as
Begin
	Declare @xKerja nvarchar(2),@xKeg nvarchar(2),@xHSatuan FLOAT,@xNilai bigint
	Declare C_HSatuan Cursor For 
	Select KD_PEKERJAAN,KD_KEGIATAN FROM VOL_RESOURCE 
	WHERE KD_PEKERJAAN NOT IN ('21','22','23','24')
	OPEN C_HSatuan
	FETCH C_HSatuan INTO @xKerja,@xKeg
	WHILE @@FETCH_STATUS = 0
	BEGIN
	   Select @xHSatuan =(Sum((A.VOL_RESOURCE)*B.HRG_RESOURCE)) 
	   From VOL_RESOURCE A,HRG_RESOURCE B
	   Where A.KD_GROUP_RESOURCE=B.KD_GROUP_RESOURCE AND
			 A.KD_RESOURCE=B.KD_RESOURCE AND
			 A.KD_PEKERJAAN =@xKerja AND
			 A.KD_KEGIATAN =@xKeg AND
			 B.THN_HRG_RESOURCE =@xTahun 
	   IF @xHSatuan IS NULL 
	   BEGIN
		SET @xHSatuan=0
	   END
	   --CEK APAKAH RECORD ADA ATAU TIDAK
	   SELECT @xNilai =COUNT(*) FROM HRG_SATUAN 
	   WHERE THN_HRG_SATUAN=@xTahun and KD_PEKERJAAN=@xKerja AND KD_KEGIATAN=@xKeg 
	   IF @xNilai=0 
	   Begin
			INSERT INTO HRG_SATUAN(KD_PROPINSI,KD_DATI2,THN_HRG_SATUAN,KD_PEKERJAAN,KD_KEGIATAN,HRG_SATUAN)
                   VALUES ('12','12',@xTahun,@xKerja,@xKeg,@xHSatuan) 
       End
       Else
       Begin
            UPDATE HRG_SATUAN SET HRG_SATUAN=@xHSatuan 
			WHERE THN_HRG_SATUAN  = @xTahun  AND KD_PEKERJAAN = @xKerja AND KD_KEGIATAN = @xKeg 
	   End
	   FETCH NEXT FROM C_HSatuan INTO @xKerja,@xKeg
	END
	CLOSE C_HSatuan
	DEALLOCATE C_HSatuan
if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
End
