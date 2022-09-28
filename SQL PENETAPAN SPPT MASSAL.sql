USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'iSPPT_MASSAL')
BEGIN
    DROP PROC iSPPT_MASSAL
END
GO
CREATE proc iSPPT_MASSAL
@xxKec1 nVarchar(3), @xxKel1 nVarchar(3),@xTahun nvarchar(4),
@xAwal_T nvarchar(4),@xAwal_B nvarchar(4),@xxJTempo datetime,@xKurang float,
@xStatus1 nvarchar(1),@xStatus2 nvarchar(1),@xStatus3 nvarchar(1),@xxTerbit datetime,@xxCetak datetime,@xNIP1 nvarchar(50),
@xSiklus smallint,@xKanwil nvarchar(2),@xKPPBB nvarchar(2),@xTunggal nvarchar(2),@Persepsi nvarchar(2),@xTP nvarchar(2),@xProses nchar(1),@xPro nvarchar(1)
as
begin


Declare @xSubjek nvarchar(30),@xProp nVarchar(2), @xKab nVarchar(2), @xxKec nVarchar(3), @xxKel nVarchar(3),
@xxBlok nVarchar(3), @xxUrut nVarchar(4), @xxJenis nVarchar(1),
@xNamaWP nvarchar(30),@xAlamatWP nvarchar(30),@xKav nvarchar(15), @xRW nvarchar(3), @xRT nvarchar(3),
@xKel1 nvarchar(30),@xKota nvarchar(30),@xPos nvarchar(5),@xNPWP nvarchar(15),@xPersil nvarchar(5),
@xKelas_T nvarchar(3),@xKelas_B nvarchar(3),
@xLuas_T float,@xLuas_B float,@xNJOP_T float,@xNJOP_B float,@xTotal float,
@xNJOPTKP int,@xNJKP float,@xHutang float,@xBayar float,@xTarif float,
@y1 nVarchar(2),@y2 nVarchar(2),@X1 nVarchar(3), @X2 nVarchar(3),@X3 nVarchar(3), @X4 nVarchar(4), @X5 nVarchar(1),
@xBumi nvarchar(1),@nMin float

if @xPro ='1'
BEGIN
	DELETE From SPPT where THN_PAJAK_SPPT=@xTahun
	--SET NOCOUNT ON;
	DECLARE C_PENETAPAN CURSOR FOR
		SELECT SUBJEK_PAJAK_ID,KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,--THN_PAJAK_SPPT,
				NM_WP,JALAN_WP,BLOK_KAV_NO_WP,RW_WP,RT_WP,KELURAHAN_WP,KOTA_WP,KD_POS_WP,NPWP,NO_PERSIL,
				TOTAL_LUAS_BUMI,TOTAL_LUAS_BNG,NJOP_BUMI,NJOP_BNG,NJOP_BUMI + NJOP_BNG as TOTAL,JNS_BUMI
		FROM QOBJEKPAJAK 
	OPEN C_PENETAPAN
		FETCH FROM C_PENETAPAN INTO @xSubjek,@xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis,--@xTahun
			@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
			@xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,@xTotal,@xBumi
		WHILE @@FETCH_STATUS = 0
		BEGIN		
			--BESARAN NJOPTKP
			Select @xNJOPTKP =NJOPTKP,@xTarif=NILAI_TARIF  FROM TARIF WHERE (NJOP_MIN*1000<=@xNJOP_B and NJOP_MAX*1000>=@xNJOP_B)
			--PILIH OBJEK YANG MENDAPATKAN NJOPTKP
			
			--SET NOCOUNT ON;
			DECLARE C_NONPAJAK CURSOR FOR
			SELECT KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP FROM DAT_SUBJEK_PAJAK_NJOPTKP where SUBJEK_PAJAK_ID=@xSubjek and THN_NJOPTKP=@xTahun
			OPEN C_NONPAJAK
			FETCH FROM C_NONPAJAK INTO @y1,@y2,@x1,@x2,@x3,@x4,@x5
			While @@FETCH_STATUS =0
			BEGIN
				if (@y1=@xProp  AND @y2=@xKab  AND @x1=@xxKec AND @x2=@xxKel AND @x3=@xxBlok AND @x4=@xxUrut  AND @x5=@xxJenis)
				begin
					set @xNJOPTKP=@xNJOPTKP
					Goto Lompat
				end
				FETCH NEXT FROM C_NONPAJAK INTO  @y1,@y2,@x1,@x2,@x3,@x4,@x5
			END
			set @xNJOPTKP=0
			Lompat:
			CLOSE C_NONPAJAK 
			DEALLOCATE C_NONPAJAK 
				--Kelas Bumi
				if @xLuas_T  =0
				Begin
					Select @xKelas_T=KD_KLS_TANAH FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH=@xAwal_T and (NILAI_MIN_TANAH*1000 <=@xNJOP_T and NILAI_MAX_TANAH*1000 >=@xNJOP_T)
				End
				Else
				Begin
					Select @xKelas_T=KD_KLS_TANAH FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH=@xAwal_T and (NILAI_MIN_TANAH*1000 <=@xNJOP_T/@xLuas_T and NILAI_MAX_TANAH*1000 >=@xNJOP_T/@xLuas_T )
				End				
				--Kelas Bangunan
				if @xLuas_B =0
				Begin
					Select @xKelas_B =KD_KLS_BNG FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG =@xAwal_B  and NILAI_MIN_BNG *1000 <=@xNJOP_B   and NILAI_MAX_BNG *1000 >=@xNJOP_B 
				End
				Else
				Begin
					Select @xKelas_B =KD_KLS_BNG FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG =@xAwal_B  and NILAI_MIN_BNG *1000 <=@xNJOP_B /@xLuas_B  and NILAI_MAX_BNG *1000 >=@xNJOP_B /@xLuas_B
				End
			IF @xNJOP_B <=0 or @xLuas_B <=0 
				Begin
					set @xKelas_B='000'
				End
			IF @xNJOP_T  <=0 or @xLuas_T  <=0 
				Begin
					set @xKelas_T ='000'
				End
			SET @xNJKP =@xTotal - @xNJOPTKP 
			if @xNJKP<0
			Begin
				Set @xNJKP=0
			End
			Set @xHutang=@xNJKP * @xTarif /100
			set @xBayar = @xHutang-@xKurang 
			Select @nMin=NILAI_PBB_MINIMAL from PBB_MINIMAL WHERE THN_PBB_MINIMAL=@xTahun
			if @xBayar<=@nMin
			Begin
				set @xBayar=@nMin
			End 
			IF @xBumi <> '4'
			Begin
			INSERT INTO SPPT(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,THN_PAJAK_SPPT,
						NM_WP_SPPT,JLN_WP_SPPT,BLOK_KAV_NO_WP_SPPT,RW_WP_SPPT,RT_WP_SPPT,KELURAHAN_WP_SPPT,KOTA_WP_SPPT,KD_POS_WP_SPPT,NPWP_SPPT,NO_PERSIL_SPPT,
						KD_KLS_TANAH,THN_AWAL_KLS_TANAH,KD_KLS_BNG,THN_AWAL_KLS_BNG,TGL_JATUH_TEMPO_SPPT,LUAS_BUMI_SPPT,LUAS_BNG_SPPT, NJOP_BUMI_SPPT,NJOP_BNG_SPPT,
						NJOP_SPPT,NJOPTKP_SPPT,NJKP_SPPT,PBB_TERHUTANG_SPPT,FAKTOR_PENGURANG_SPPT,PBB_YG_HARUS_DIBAYAR_SPPT,
						STATUS_PEMBAYARAN_SPPT,STATUS_TAGIHAN_SPPT,STATUS_CETAK_SPPT,TGL_TERBIT_SPPT,TGL_CETAK_SPPT,NIP_PENCETAK_SPPT,SIKLUS_SPPT,KD_KANWIL_BANK,KD_KPPBB_BANK,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,KD_TP,PROSES)
			Values(@xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis, @xTahun,
					@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
					@xKelas_T,@xAwal_T,@xKelas_B,@xAwal_B,@xxJTempo, @xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,
					@xTotal,@xNJOPTKP,@xNJKP,@xHutang,@xKurang,round(@xBayar,0),
					@xStatus1,@xStatus2,@xStatus3,@xxTerbit,@xxCetak,@xNIP1,@xSiklus,@xKanwil,@xKPPBB,@xTunggal,@Persepsi,@xTP,@xProses)
			END		
			FETCH NEXT FROM C_PENETAPAN INTO  @xSubjek,@xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis,--@xTahun
			@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
			@xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,@xTotal,@xBumi
		END
		CLOSE C_PENETAPAN
		DEALLOCATE C_PENETAPAN
	END	
ELSE
	IF @xPro = '2'
	BEGIN
	DELETE From SPPT where KD_KECAMATAN=@xxKec1 and THN_PAJAK_SPPT=@xTahun
	SET NOCOUNT ON;
	DECLARE C_PENETAPAN CURSOR FOR
		SELECT SUBJEK_PAJAK_ID,KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,--THN_PAJAK_SPPT,
				NM_WP,JALAN_WP,BLOK_KAV_NO_WP,RW_WP,RT_WP,KELURAHAN_WP,KOTA_WP,KD_POS_WP,NPWP,NO_PERSIL,
				TOTAL_LUAS_BUMI,TOTAL_LUAS_BNG,NJOP_BUMI,NJOP_BNG,NJOP_BUMI + NJOP_BNG as TOTAL,JNS_BUMI
		FROM QOBJEKPAJAK WHERE KD_KECAMATAN=@xxKec1 
	OPEN C_PENETAPAN
		FETCH FROM C_PENETAPAN INTO @xSubjek,@xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis,--@xTahun
			@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
			@xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,@xTotal,@xBumi
		WHILE @@FETCH_STATUS = 0
		BEGIN		
			--BESARAN NJOPTKP
			Select @xNJOPTKP =NJOPTKP,@xTarif=NILAI_TARIF  FROM TARIF WHERE (NJOP_MIN*1000<=@xNJOP_B and NJOP_MAX*1000>=@xNJOP_B)
			--PILIH OBJEK YANG MENDAPATKAN NJOPTKP
			
			--SET NOCOUNT ON;
			DECLARE C_NONPAJAK CURSOR FOR
			SELECT KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP FROM DAT_SUBJEK_PAJAK_NJOPTKP where SUBJEK_PAJAK_ID=@xSubjek and KD_KECAMATAN=@xxKec1 and THN_NJOPTKP=@xTahun
			OPEN C_NONPAJAK
			FETCH FROM C_NONPAJAK INTO @y1,@y2,@x1,@x2,@x3,@x4,@x5
			While @@FETCH_STATUS =0
			BEGIN
				if (@y1=@xProp  AND @y2=@xKab  AND @x1=@xxKec AND @x2=@xxKel AND @x3=@xxBlok AND @x4=@xxUrut  AND @x5=@xxJenis)
				begin
					set @xNJOPTKP=@xNJOPTKP
					Goto KeluaR1
				end
				FETCH NEXT FROM C_NONPAJAK INTO  @y1,@y2,@x1,@x2,@x3,@x4,@x5
			END
			set @xNJOPTKP=0
			KeluaR1:
			CLOSE C_NONPAJAK 
			DEALLOCATE C_NONPAJAK 
				--Kelas Bumi
				if @xLuas_T  =0
				Begin
					Select @xKelas_T=KD_KLS_TANAH FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH=@xAwal_T and (NILAI_MIN_TANAH*1000 <=@xNJOP_T and NILAI_MAX_TANAH*1000 >=@xNJOP_T)
				End
				Else
				Begin
					Select @xKelas_T=KD_KLS_TANAH FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH=@xAwal_T and (NILAI_MIN_TANAH*1000 <=@xNJOP_T/@xLuas_T and NILAI_MAX_TANAH*1000 >=@xNJOP_T/@xLuas_T )
				End				
				--Kelas Bangunan
				if @xLuas_B =0
				Begin
					Select @xKelas_B =KD_KLS_BNG FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG =@xAwal_B  and NILAI_MIN_BNG *1000 <=@xNJOP_B   and NILAI_MAX_BNG *1000 >=@xNJOP_B 
				End
				Else
				Begin
					Select @xKelas_B =KD_KLS_BNG FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG =@xAwal_B  and NILAI_MIN_BNG *1000 <=@xNJOP_B /@xLuas_B  and NILAI_MAX_BNG *1000 >=@xNJOP_B /@xLuas_B
				End
			IF @xNJOP_B <=0 or @xLuas_B <=0 
				Begin
					set @xKelas_B='000'
				End
			IF @xNJOP_T  <=0 or @xLuas_T  <=0 
				Begin
					set @xKelas_T ='000'
				End
			SET @xNJKP =@xTotal - @xNJOPTKP 
			if @xNJKP<0
			Begin
				Set @xNJKP=0
			End
			Set @xHutang=@xNJKP * @xTarif /100
			set @xBayar = @xHutang-@xKurang 
			Select @nMin=NILAI_PBB_MINIMAL from PBB_MINIMAL WHERE THN_PBB_MINIMAL=@xTahun
			if @xBayar<=@nMin
			Begin
				set @xBayar=@nMin
			End 
			IF @xBumi <> '4'
			Begin
			INSERT INTO SPPT(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,THN_PAJAK_SPPT,
						NM_WP_SPPT,JLN_WP_SPPT,BLOK_KAV_NO_WP_SPPT,RW_WP_SPPT,RT_WP_SPPT,KELURAHAN_WP_SPPT,KOTA_WP_SPPT,KD_POS_WP_SPPT,NPWP_SPPT,NO_PERSIL_SPPT,
						KD_KLS_TANAH,THN_AWAL_KLS_TANAH,KD_KLS_BNG,THN_AWAL_KLS_BNG,TGL_JATUH_TEMPO_SPPT,LUAS_BUMI_SPPT,LUAS_BNG_SPPT, NJOP_BUMI_SPPT,NJOP_BNG_SPPT,
						NJOP_SPPT,NJOPTKP_SPPT,NJKP_SPPT,PBB_TERHUTANG_SPPT,FAKTOR_PENGURANG_SPPT,PBB_YG_HARUS_DIBAYAR_SPPT,
						STATUS_PEMBAYARAN_SPPT,STATUS_TAGIHAN_SPPT,STATUS_CETAK_SPPT,TGL_TERBIT_SPPT,TGL_CETAK_SPPT,NIP_PENCETAK_SPPT,SIKLUS_SPPT,KD_KANWIL_BANK,KD_KPPBB_BANK,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,KD_TP,PROSES)
			Values(@xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis, @xTahun,
					@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
					@xKelas_T,@xAwal_T,@xKelas_B,@xAwal_B,@xxJTempo, @xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,
					@xTotal,@xNJOPTKP,@xNJKP,@xHutang,@xKurang,round(@xBayar,0),
					@xStatus1,@xStatus2,@xStatus3,@xxTerbit,@xxCetak,@xNIP1,@xSiklus,@xKanwil,@xKPPBB,@xTunggal,@Persepsi,@xTP,@xProses)
			END		
			FETCH NEXT FROM C_PENETAPAN INTO  @xSubjek,@xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis,--@xTahun
			@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
			@xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,@xTotal,@xBumi
		END
		CLOSE C_PENETAPAN
		DEALLOCATE C_PENETAPAN
	END

ELSE
	--SET NOCOUNT ON;
	if @xPro = '3'
	BEGIN
	DELETE  From SPPT where KD_KECAMATAN=@xxKec1 and KD_KELURAHAN=@xxKel1 and THN_PAJAK_SPPT=@xTahun
	SET NOCOUNT ON;
	DECLARE C_PENETAPAN CURSOR FOR
		SELECT SUBJEK_PAJAK_ID,KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,--THN_PAJAK_SPPT,
				NM_WP,JALAN_WP,BLOK_KAV_NO_WP,RW_WP,RT_WP,KELURAHAN_WP,KOTA_WP,KD_POS_WP,NPWP,NO_PERSIL,
				TOTAL_LUAS_BUMI,TOTAL_LUAS_BNG,NJOP_BUMI,NJOP_BNG,NJOP_BUMI + NJOP_BNG as TOTAL,JNS_BUMI
		FROM QOBJEKPAJAK WHERE KD_KECAMATAN=@xxKec1 and KD_KELURAHAN=@xxKel1
	OPEN C_PENETAPAN
		FETCH FROM C_PENETAPAN INTO @xSubjek, @xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis,--@xTahun
			@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
			@xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,@xTotal,@xBumi
		WHILE @@FETCH_STATUS = 0
		BEGIN
			
			--BESARAN NJOPTKP
			Select @xNJOPTKP =NJOPTKP,@xTarif=NILAI_TARIF  FROM TARIF WHERE (NJOP_MIN*1000<=@xNJOP_B and NJOP_MAX*1000>=@xNJOP_B)
			--PILIH OBJEK YANG MENDAPATKAN NJOPTKP
--			SET NOCOUNT OFF;
			SET NOCOUNT ON;
			DECLARE C_NONPAJAK CURSOR FOR
			SELECT KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP FROM DAT_SUBJEK_PAJAK_NJOPTKP where SUBJEK_PAJAK_ID=@xSubjek and KD_KECAMATAN=@xxKec1 and KD_KELURAHAN=@xxKel1 and THN_NJOPTKP=@xTahun
			OPEN C_NONPAJAK
			FETCH FROM C_NONPAJAK INTO @y1,@y2,@x1,@x2,@x3,@x4,@x5
			While @@FETCH_STATUS =0
			BEGIN
				if (@y1=@xProp  AND @y2=@xKab  AND @x1=@xxKec AND @x2=@xxKel AND @x3=@xxBlok AND @x4=@xxUrut  AND @x5=@xxJenis)
				begin
					set @xNJOPTKP=@xNJOPTKP
					goto Keluar2
				end
				FETCH NEXT FROM C_NONPAJAK INTO  @y1,@y2,@x1,@x2,@x3,@x4,@x5
			END
			set @xNJOPTKP=0
			Keluar2:
			CLOSE C_NONPAJAK 
			DEALLOCATE C_NONPAJAK 
				--Kelas Bumi
				if @xLuas_T  =0
				Begin
					Select @xKelas_T=KD_KLS_TANAH FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH=@xAwal_T and (NILAI_MIN_TANAH*1000 <=@xNJOP_T and NILAI_MAX_TANAH*1000 >=@xNJOP_T)
				End
				Else
				Begin
					Select @xKelas_T=KD_KLS_TANAH FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH=@xAwal_T and (NILAI_MIN_TANAH*1000 <=@xNJOP_T/@xLuas_T and NILAI_MAX_TANAH*1000 >=@xNJOP_T/@xLuas_T )
				End
				
				--Kelas Bangunan
				if @xLuas_B =0
				Begin
					Select @xKelas_B =KD_KLS_BNG FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG =@xAwal_B  and NILAI_MIN_BNG *1000 <=@xNJOP_B   and NILAI_MAX_BNG *1000 >=@xNJOP_B 
				End
				Else
				Begin
					Select @xKelas_B =KD_KLS_BNG FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG =@xAwal_B  and NILAI_MIN_BNG *1000 <=@xNJOP_B /@xLuas_B  and NILAI_MAX_BNG *1000 >=@xNJOP_B /@xLuas_B
				End
			IF @xNJOP_B <=0 or @xLuas_B <=0 
				Begin
					set @xKelas_B='000'
				End
			IF @xNJOP_T  <=0 or @xLuas_T  <=0 
				Begin
					set @xKelas_T ='000'
				End
			SET @xNJKP =@xTotal - @xNJOPTKP 
			if @xNJKP<0
			Begin
				Set @xNJKP=0
			End
			Set @xHutang=@xNJKP * @xTarif /100
			set @xBayar = @xHutang-@xKurang 
			Select @nMin=NILAI_PBB_MINIMAL from PBB_MINIMAL WHERE THN_PBB_MINIMAL=@xTahun
			if @xBayar<=@nMin
			Begin
				set @xBayar=@nMin
			End 
			IF @xBumi <> '4'
			Begin
			INSERT INTO SPPT(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,THN_PAJAK_SPPT,
						NM_WP_SPPT,JLN_WP_SPPT,BLOK_KAV_NO_WP_SPPT,RW_WP_SPPT,RT_WP_SPPT,KELURAHAN_WP_SPPT,KOTA_WP_SPPT,KD_POS_WP_SPPT,NPWP_SPPT,NO_PERSIL_SPPT,
						KD_KLS_TANAH,THN_AWAL_KLS_TANAH,KD_KLS_BNG,THN_AWAL_KLS_BNG,TGL_JATUH_TEMPO_SPPT,LUAS_BUMI_SPPT,LUAS_BNG_SPPT, NJOP_BUMI_SPPT,NJOP_BNG_SPPT,
						NJOP_SPPT,NJOPTKP_SPPT,NJKP_SPPT,PBB_TERHUTANG_SPPT,FAKTOR_PENGURANG_SPPT,PBB_YG_HARUS_DIBAYAR_SPPT,
						STATUS_PEMBAYARAN_SPPT,STATUS_TAGIHAN_SPPT,STATUS_CETAK_SPPT,TGL_TERBIT_SPPT,TGL_CETAK_SPPT,NIP_PENCETAK_SPPT,SIKLUS_SPPT,KD_KANWIL_BANK,KD_KPPBB_BANK,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,KD_TP,PROSES)
			Values(@xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis, @xTahun,
					@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
					@xKelas_T,@xAwal_T,@xKelas_B,@xAwal_B,@xxJTempo, @xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,
					@xTotal,@xNJOPTKP,@xNJKP,@xHutang,@xKurang,round(@xBayar,0),
					@xStatus1,@xStatus2,@xStatus3,@xxTerbit,@xxCetak,@xNIP1,@xSiklus,@xKanwil,@xKPPBB,@xTunggal,@Persepsi,@xTP,@xProses)
			END		
			FETCH NEXT FROM C_PENETAPAN INTO  @xSubjek, @xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis,--@xTahun
			@xNamaWP,@xAlamatWP,@xKav, @xRW, @xRT,@xKel1,@xKota,@xPos,@xNPWP,@xPersil,
			@xLuas_T,@xLuas_B,@xNJOP_T,@xNJOP_B,@xTotal,@xBumi
		END
		CLOSE C_PENETAPAN
		DEALLOCATE C_PENETAPAN
	END
SET NOCOUNT ON;	
if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
end