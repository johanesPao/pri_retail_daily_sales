class Kueri:
    def __init__(self, parameter: dict[str, any]):
        self.parameter = {
            "bc": {
                "tabel": parameter["bc"]["tabel"],
                "kolom": parameter["bc"]["kolom"],
                "argumen": parameter["bc"]["argumen"]
            },
            "pd": {
                "nama": parameter["pd"]["nama"],
                "tabel": parameter["pd"]["tabel"],
                "kolom": parameter["pd"]["kolom"]
            }
        }

    def kueri_target(self, tgl_report: str):
        return f"""
            with store as (
                select distinct
                    kode_toko."{self.parameter["bc"]["kolom"]["loc_code"]}" "Toko",
                    case
                        when kode_toko."{self.parameter["bc"]["kolom"]["loc_code"]}" like 'BZ%%'
                        then concat('BAZAAR ', right(kode_toko."{self.parameter["bc"]["kolom"]["loc_code"]}", 3))
                        else nama_toko."{self.parameter["bc"]["kolom"]["name"]}"
                    end "Nama Toko"
                from "{self.parameter["bc"]["tabel"]["pri"] + self.parameter["bc"]["tabel"]["store_map_5ec"]}" kode_toko
                left join "{self.parameter["bc"]["tabel"]["pri"] + self.parameter["bc"]["tabel"]["dim_val_437"]}" nama_toko
                on
                    kode_toko."{self.parameter["bc"]["kolom"]["store_dim"]}" = nama_toko."{self.parameter["bc"]["kolom"]["code"]}"
                where
                    kode_toko."{self.parameter["bc"]["kolom"]["store_dim"]}" != ''
            ), d as (
                select distinct
                    "{self.parameter["bc"]["kolom"]["store_no"]}" "Toko",
                    SUM("{self.parameter["bc"]["kolom"]["net_amt"]}" * -1) over(partition by "{self.parameter["bc"]["kolom"]["store_no"]}") "Daily Sales"
                from "{self.parameter["bc"]["tabel"]["pri"] + self.parameter["bc"]["tabel"]["store_sales_entry_5ec"]}"
                where 
                    "{self.parameter["bc"]["kolom"]["date"]}" = '{tgl_report}'
            ), 
            -- limit sales by week according to date, year and cut the data
            -- based on month. Like in the case of week 48 2023 is 27 Nov -
            -- 3 Dec, but we want to sum only 1 - 3 Dec.
            -- Not sure why you want this. This may introduce a misinformation, since
            -- we're going to have 2 week 48 (27-30 Nov and 1-3 Dec), and to compare it
            -- to wtd previous year? Does week 48 in the previous year for month of
            -- december also span from 1 to 3? Certainly not.
            -- I think if you really want to analyze weekly sales, don't cut it by month
            -- and follow the ISO 8601 convention (https://en.wikipedia.org/wiki/ISO_week_date)
            -- UPDATE: We're going to go week of the year vs week of the year comparison, no need
            -- to limit the data for specific month, since week of the year will capture the daily
            -- sales pattern already
            wtd as (
                select distinct
                "{self.parameter["bc"]["kolom"]["store_no"]}" "Toko",
                SUM("{self.parameter["bc"]["kolom"]["net_amt"]}" * -1) over(partition by "{self.parameter["bc"]["kolom"]["store_no"]}") "WTD Sales"
                from "{self.parameter["bc"]["tabel"]["pri"] + self.parameter["bc"]["tabel"]["store_sales_entry_5ec"]}"
                where
                date_part('year', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('year', '{tgl_report}'::Date) and
                date_part('week', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('week', '{tgl_report}'::Date) and
                -- dalam kasus laporan ini ditarik di masa depan, limit to date sampai dengan tanggal
                -- laporan ini
                "{self.parameter["bc"]["kolom"]["date"]}" <= '{tgl_report}'
            ), wtd_ly as (
                select distinct
                "{self.parameter["bc"]["kolom"]["store_no"]}" "Toko",
                SUM("{self.parameter["bc"]["kolom"]["net_amt"]}" * -1) over(partition by "{self.parameter["bc"]["kolom"]["store_no"]}") "WTD LY Sales"
                from "{self.parameter["bc"]["tabel"]["pri"] + self.parameter["bc"]["tabel"]["store_sales_entry_5ec"]}"
                where
                date_part('year', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('year', '{tgl_report}'::Date) -1 and	-- -1 for last year
                date_part('week', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('week', '{tgl_report}'::Date) and		-- same week
                date_part('dow', "{self.parameter["bc"]["kolom"]["date"]}") in (
                    select distinct
                    date_part('dow', "{self.parameter["bc"]["kolom"]["date"]}")
                    from "{self.parameter["bc"]["tabel"]["pri"] + self.parameter["bc"]["tabel"]["store_sales_entry_5ec"]}"
                    where
                    date_part('year', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('year', '{tgl_report}'::Date)  and		-- this year
                    date_part('week', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('week', '{tgl_report}'::Date) and		-- same week
                    "{self.parameter["bc"]["kolom"]["date"]}" <= '{tgl_report}'
                )
            ), mtd as (
                select distinct
                "{self.parameter["bc"]["kolom"]["store_no"]}" "Toko",
                SUM("{self.parameter["bc"]["kolom"]["net_amt"]}" * -1) over(partition by "{self.parameter["bc"]["kolom"]["store_no"]}") "MTD Sales"
                from "{self.parameter["bc"]["tabel"]["pri"] + self.parameter["bc"]["tabel"]["store_sales_entry_5ec"]}"
                where
                date_part('year', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('year', '{tgl_report}'::Date) and
                date_part('month', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('month', '{tgl_report}'::Date) and 
                "{self.parameter["bc"]["kolom"]["date"]}" <= '{tgl_report}'
            ), mtd_ly as (
                select distinct
                "{self.parameter["bc"]["kolom"]["store_no"]}" "Toko",
                SUM("{self.parameter["bc"]["kolom"]["net_amt"]}" * -1) over(partition by "{self.parameter["bc"]["kolom"]["store_no"]}") "MTD LY Sales"
                from "{self.parameter["bc"]["tabel"]["pri"] + self.parameter["bc"]["tabel"]["store_sales_entry_5ec"]}"
                where
                date_part('year', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('year', '{tgl_report}'::Date) - 1 and			-- -1 for last year
                date_part('month', "{self.parameter["bc"]["kolom"]["date"]}") = date_part('month', '{tgl_report}'::Date) and				-- same month
                date_part('day', "{self.parameter["bc"]["kolom"]["date"]}") <= date_part('day', '{tgl_report}'::Date)					-- day less than or equal to
            )
            select distinct
                store."Toko",
                store."Nama Toko",
                d."Daily Sales",
                wtd."WTD Sales",
                wtd_ly."WTD LY Sales",
                mtd."MTD Sales",
                mtd_ly."MTD LY Sales"
            from store
            left join d
            on
                store."Toko" = d."Toko"
            left join wtd
            on
                store."Toko" = wtd."Toko"
            left join mtd
            on
                store."Toko" = mtd."Toko"
            left join wtd_ly
            on
                store."Toko" = wtd_ly."Toko"
            left join mtd_ly
            on
                store."Toko" = mtd_ly."Toko"
            order by store."Toko"
        """


    def kueri_sales(self, tgl_report: str):
        return f"""
            with toko AS (
                select distinct
                    "{self.parameter["pd"]["kolom"]["kt"]}" "Toko"
                from {self.parameter["pd"]["nama"]}."{self.parameter["pd"]["tabel"]["ms"]}"
            ), d AS (
                select distinct
                    "{self.parameter["pd"]["kolom"]["kt"]}" "Toko",
                    sum("{self.parameter["pd"]["kolom"]["ntnp"]}") over(partition by "{self.parameter["pd"]["kolom"]["kt"]}") "Daily Target"
                from {self.parameter["pd"]["nama"]}."{self.parameter["pd"]["tabel"]["dt"]}"
                where
                    "{self.parameter["pd"]["kolom"]["t"]}" = '{tgl_report}'
            ), wtd AS (
                select distinct
                    "{self.parameter["pd"]["kolom"]["kt"]}" "Toko",
                    sum("{self.parameter["pd"]["kolom"]["ntnp"]}") over(partition by "{self.parameter["pd"]["kolom"]["kt"]}") "WTD Target"
                from {self.parameter["pd"]["nama"]}."{self.parameter["pd"]["tabel"]["dt"]}"
                where
                    date_part('year', "{self.parameter["pd"]["kolom"]["t"]}") = date_part('year', '{tgl_report}'::Date) and
                    date_part('week', "{self.parameter["pd"]["kolom"]["t"]}") = date_part('week', '{tgl_report}'::Date) and
                    "{self.parameter["pd"]["kolom"]["t"]}" <= '{tgl_report}'
            ), mtd AS (
                select distinct
                    "{self.parameter["pd"]["kolom"]["kt"]}" "Toko",
                    sum("{self.parameter["pd"]["kolom"]["ntnp"]}") over(partition by "{self.parameter["pd"]["kolom"]["kt"]}") "MTD Target"
                from {self.parameter["pd"]["nama"]}."{self.parameter["pd"]["tabel"]["dt"]}"
                where
                    date_part('year', "{self.parameter["pd"]["kolom"]["t"]}") = date_part('year', '{tgl_report}'::Date) and
                    date_part('month', "{self.parameter["pd"]["kolom"]["t"]}") = date_part('month', '{tgl_report}'::Date) and
                    "{self.parameter["pd"]["kolom"]["t"]}" <= '{tgl_report}'
            )
            select
                toko."Toko",
                case
                    when d."Daily Target" is null
                    then 0
                    else d."Daily Target"
                end "Daily Target",
                case
                    when wtd."WTD Target" is null
                    then 0
                    else wtd."WTD Target"
                end "WTD Target",
                case
                    when mtd."MTD Target" is null
                    then 0
                    else mtd."MTD Target"
                end "MTD Target"
            from toko
            left join d
            on
                toko."Toko" = d."Toko"
            left join wtd
            on
                toko."Toko" = wtd."Toko"
            left join mtd
            on
                toko."Toko" = mtd."Toko"
        """


    def kueri_area_toko(self, tgl_report: str) -> str:
        return f"""
            select distinct
                at."{self.parameter["pd"]["kolom"]["kt"]}" "Toko",
                max_tanggal_efektif."{self.parameter["pd"]["kolom"]["a"]}" "Area"
            from {self.parameter["pd"]["nama"]}."{self.parameter["pd"]["tabel"]["at"]}" at
            left join (
                select distinct
                    "{self.parameter["pd"]["kolom"]["kt"]}",
                    "{self.parameter["pd"]["kolom"]["a"]}",
                    max("{self.parameter["pd"]["kolom"]["te"]}") over(partition by "{self.parameter["pd"]["kolom"]["kt"]}", "{self.parameter["pd"]["kolom"]["a"]}") "max_tanggal_efektif"
                from {self.parameter["pd"]["nama"]}."{self.parameter["pd"]["tabel"]["at"]}"
                where
                    "{self.parameter["pd"]["kolom"]["te"]}"::Date <= '{tgl_report}'
            ) max_tanggal_efektif
            on
                at."{self.parameter["pd"]["kolom"]["kt"]}" = max_tanggal_efektif."{self.parameter["pd"]["kolom"]["kt"]}" and
                at."{self.parameter["pd"]["kolom"]["te"]}" = max_tanggal_efektif."max_tanggal_efektif"
        """
