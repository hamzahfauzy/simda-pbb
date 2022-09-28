USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'iSPPT_TUNGGAL')
BEGIN
    DROP PROC iSPPT_TUNGGAL
END
GO
CREATE proc iSPPT_TUNGGAL
@xProp nVarchar(2), @xKab nVarchar(2), @xxKec nVarchar(3), @xxKel nVarchar(3),
@xxBlok nVarchar(3), @xxUrut nVarchar(4), @xxJenis nVarchar(1),@xTahun nvarchar(4),
@xNamaWP nvarchar(30),@xAlamatWP nvarchar(30),@xKav nvarchar(15), @xRW nvarchar(3), @xRT nvarchar(3),
@xKel1 nvarchar(30),@xKota nvarchar(30),@xPos nvarchar(5),@xNPWP nvarchar(15),@xPersil nvarchar(5),
@xKelas_T nvarchar(3),@xAwal_T nvarchar(4),@xKelas_B nvarchar(3),@xAwal_B nvarchar(4),@xxJTempo datetime, 
@xLuas_T float,@xLuas_B float,@xNJOP_T float,@xNJOP_B float,@xTotal float,@xNJOPTKP int,@xNJKP float,@xHutang float,@xKurang float,@xBayar float,
@xStatus1 nvarchar(1),@xStatus2 nvarchar(1),@xStatus3 nvarchar(1),@xxTerbit datetime,@xxCetak datetime,@xNIP1 nvarchar(50),
@xSiklus smallint,@xKanwil nvarchar(2),@xKPPBB nvarchar(2),@xTunggal nvarchar(2),@Persepsi nvarchar(2),@xTP nvarchar(2),@xProses nchar(1),
@xCPP nvarchar(1),@xNOP nvarchar(30)
as
begin
Delete From SPPT where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)=@xNOP and THN_PAJAK_SPPT=@xTahun AND PROSES=@xCPP 

INSERT INTO SPPT(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,THN_PAJAK_SPPT,
			NM_WP_SPPT,JLN_WP_SPPT,BLOK_KAV_NO_WP_SPPT,RW_WP_SPPT,RT_WP_SPPT,KELURAHAN_WP_SPPT,KOTA_WP_SPPT,KD_POS_WP_SPPT,NPWP_SPPT,NO_PERSIL_SPPT,
			KD_KLS_TANAH,THN_AWAL_KLS_TANAH,KD_KLS_BNG,THN_AWAL_KLS_BNG,TGL_JATUH_TEMPO_SPPT,LUAS_BUMI_SPPT,LUAS_BNG_SPPT, NJOP_BUMI_SPPT,NJOP_BNG_SPPT,
			NJOP_SPPT,NJOPTKP_SPPT,NJKP_SPPT,PBB_TERHUTANG_SPPT,FAKTOR_PENGURANG_SPPT,PBB_YG_HARUS_DIBAYAR_SPPT,
			STATUS_PEMBAYARAN_SPPT,STATUS_TAGIHAN_SPPT,STATUS_CETAK_SPPT,TGL_TERBIT_SPPT,TGL_CETAK_SPPT,NIP_PENCETAK_SPPT,SIKLUS_SPPT,KD_KANWIL_BANK,KD_KPPBB_BANK,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,KD_TP,PROSES)
Values(@xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis, @xTahun,
		@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
		@xKelas_T,@xAwal_T,@xKelas_B,@xAwal_B,@xxJTempo, @xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,
		@xTotal,@xNJOPTKP,@xNJKP,@xHutang,@xKurang,@xBayar,
		@xStatus1,@xStatus2,@xStatus3,@xxTerbit,@xxCetak,@xNIP1,@xSiklus,@xKanwil,@xKPPBB,@xTunggal,@Persepsi,@xTP,@xProses)

if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
end