USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'HAPUS_LUNAS_TUNGGAL')
BEGIN
    DROP PROC HAPUS_LUNAS_TUNGGAL
END
GO
CREATE proc HAPUS_LUNAS_TUNGGAL
@xTahun nvarchar(4),@BayarKe smallint,@xNOP nvarchar(30)
AS
BEGIN
	DELETE  From PEMBAYARAN_SPPT where KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= @xNOP and THN_PAJAK_SPPT=@xTahun AND PEMBAYARAN_SPPT_KE=@BayarKe 
	UPDATE SPPT SET STATUS_PEMBAYARAN_SPPT ='0' WHERE KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= @xNOP and (PROSES='M' OR PROSES='T')and THN_PAJAK_SPPT=@xTahun 
if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
end		