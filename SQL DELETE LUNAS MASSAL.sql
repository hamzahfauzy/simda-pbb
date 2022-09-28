USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'HAPUS_LUNAS_MASSAL')
BEGIN
    DROP PROC HAPUS_LUNAS_MASSAL
END
GO
CREATE proc [HAPUS_LUNAS_MASSAL]
@xxKec nVarchar(3), @xxKel nVarchar(3),
@xTahun nvarchar(4),@xProses nvarchar(1)
AS
BEGIN
if @xProses = '1'
begin
	DELETE  From PEMBAYARAN_SPPT WHERE THN_PAJAK_SPPT=@xTahun
	UPDATE SPPT SET STATUS_PEMBAYARAN_SPPT ='0' WHERE THN_PAJAK_SPPT=@xTahun 
end
else
if @xProses = '2'
begin
	DELETE  From PEMBAYARAN_SPPT WHERE KD_KECAMATAN=@xxKec and THN_PAJAK_SPPT=@xTahun
	UPDATE SPPT SET STATUS_PEMBAYARAN_SPPT ='0' WHERE KD_KECAMATAN=@xxKec and THN_PAJAK_SPPT=@xTahun 
end
else
if @xProses = '3'
begin
	DELETE  From PEMBAYARAN_SPPT WHERE KD_KECAMATAN=@xxKec and KD_KELURAHAN=@xxKel and THN_PAJAK_SPPT=@xTahun
	UPDATE SPPT SET STATUS_PEMBAYARAN_SPPT ='0' WHERE KD_KECAMATAN=@xxKec and KD_KELURAHAN =@xxKel and THN_PAJAK_SPPT=@xTahun 
end


if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
end		