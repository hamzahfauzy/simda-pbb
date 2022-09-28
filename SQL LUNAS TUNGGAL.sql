USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'LUNAS_TUNGGAL')
BEGIN
    DROP PROC LUNAS_TUNGGAL
END
GO
CREATE proc LUNAS_TUNGGAL
@xTahun nvarchar(4),@BayarKe smallint,@xTP nvarchar(2),--smallint,@xKanwil nvarchar(2),@xKPPBB nvarchar(2),@xTunggal nvarchar(2),@xPersepsi nvarchar(2),
@xDenda float,@TBayar datetime, @TRekam datetime, @xNIP nvarchar(30),@xNOP nvarchar(30)
 
AS
BEGIN
Declare @xBayar float
Declare @xProp nVarchar(3), @xKab nVarchar(3), @xxKec nVarchar(3), @xxKel nVarchar(3),
@xxBlok nVarchar(3), @xxUrut nVarchar(4), @xxJenis nVarchar(1)


SELECT @xProp=KD_PROPINSI,@xKab=KD_DATI2,@xxKec=KD_KECAMATAN,@xxKel=KD_KELURAHAN,@xxBlok=KD_BLOK,@xxUrut=NO_URUT,@xxJenis=KD_JNS_OP,
		--,@BayarKe=PEMBAYARAN_SPPT_KE,--@xKanwil=KD_KANWIL_BANK, @xKPPBB= KD_KPPBB_BANK,@xTunggal=KD_BANK_TUNGGAL,@xPersepsi=KD_BANK_PERSEPSI,
		@xTahun=THN_PAJAK_SPPT,@xBayar=PBB_YG_HARUS_DIBAYAR_SPPT
		FROM SPPT WHERE KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= @xNOP and (PROSES='M' OR PROSES='T')and THN_PAJAK_SPPT=@xTahun 
DELETE  From PEMBAYARAN_SPPT where KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= @xNOP and THN_PAJAK_SPPT=@xTahun 

INSERT INTO PEMBAYARAN_SPPT(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,THN_PAJAK_SPPT,PEMBAYARAN_SPPT_KE,KD_KANWIL_BANK, KD_KPPBB_BANK,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,KD_TP,DENDA_SPPT,JML_SPPT_YG_DIBAYAR,TGL_PEMBAYARAN_SPPT,TGL_REKAM_BYR_SPPT,NIP_REKAM_BYR_SPPT)
	values (@xProp,@xKab,@xxKec,@xxKel,@xxBlok,@xxUrut,@xxJenis, @xTahun,@BayarKe,'01', '16','01','01', @xTP,@xDenda,@xBayar+@xDenda,@TBayar,@TRekam,@xNIP) 	

UPDATE SPPT SET STATUS_PEMBAYARAN_SPPT ='1' WHERE KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= @xNOP and (PROSES='M' OR PROSES='T')and THN_PAJAK_SPPT=@xTahun 

if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
end		
    