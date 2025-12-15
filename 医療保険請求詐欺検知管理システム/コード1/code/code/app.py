import os
import io
import random
import re
import string
from datetime import datetime
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, send_file, abort, flash
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, login_user, login_required,
    logout_user, current_user, UserMixin
)
from werkzeug.security import generate_password_hash, check_password_hash
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import numpy as np
import joblib
from sklearn.preprocessing import MinMaxScaler, LabelEncoder

# —— Flask & DB 初期設定 (初始化设置) ——
app = Flask(__name__)
app.config.update(
    SECRET_KEY='replace-with-your-secret-key',
    SQLALCHEMY_DATABASE_URI='sqlite:///app.db',
    SQLALCHEMY_TRACK_MODIFICATIONS=False
)
db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

MODEL_PATH = 'best_xgboost_model.pkl'
DATA_PATH = '测试数据.xlsx'


# —— ORM モデル (ORM 模型) ——
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False)
    password_hash = db.Column(db.String(128), nullable=False)

    def set_password(self, pw):
        self.password_hash = generate_password_hash(pw)

    def check_password(self, pw):
        return check_password_hash(self.password_hash, pw)


class InsuranceClaim(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cl_no = db.Column(db.String, unique=True, nullable=False)
    # 2、3
    incur_date_from = db.Column(db.Date)  # 保険事故発生開始日 (出险开始日期)
    incur_date_to = db.Column(db.Date)  # 保険事故発生終了日 (出险结束日期)
    # 4
    ben_head = db.Column(db.String)  # 給付項目 (福利项目)
    # 5
    diag_code = db.Column(db.String)  # 疾病コード (疾病代码)
    # 6
    codes = db.Column(db.String)  # バーコードリスト (条形码列表)
    # 7
    prov_name = db.Column(db.String)  # 医療機関名 (医疗机构名称)
    # 8
    pay_date = db.Column(db.DateTime)  # 振込日時 (划账时间)
    # 9
    pay_amt = db.Column(db.Float)  # 支払金額 (赔付金额)
    cl_line_status = db.Column(db.String)  # ステータス: 'AC'(承認), 'PD'(保留), 'PV'(保留検証), 'RJ'(拒絶)
    prov_level = db.Column(db.Integer)
    invoice_cnt = db.Column(db.Float)
    cl_third_party_pay_amt = db.Column(db.Float)
    cwf_amt_day = db.Column(db.Float)
    codes_count = db.Column(db.Integer)
    cl_owner_pay_amt = db.Column(db.Float)
    pay_amt_usd = db.Column(db.Float)
    app_amt = db.Column(db.Float)
    ben_spend = db.Column(db.Float)
    diag_code_prefix = db.Column(db.Integer)
    ben_type = db.Column(db.Integer)
    ded_amt = db.Column(db.Float)


# —— 初回リクエスト時にテーブル作成とデフォルトユーザー追加 (首次请求时建表并插入默认用户) ——
@app.before_first_request
def init_db():
    db.create_all()
    if not User.query.first():
        u = User(username='admin')
        u.set_password('admin')
        db.session.add(u)
        db.session.commit()


@login_manager.user_loader
def load_user(uid):
    return User.query.get(int(uid))


# —— CAPTCHA画像生成 (验证码图像) ——
@app.route('/captcha.png')
def captcha():
    code = session.get('captcha_text', '')
    img = Image.new('RGB', (100, 30), (255, 255, 255))
    d = ImageDraw.Draw(img)
    f = ImageFont.load_default()
    d.text((5, 5), code, font=f, fill=(0, 0, 0))
    buf = io.BytesIO()
    img.save(buf, 'PNG')
    buf.seek(0)
    return send_file(buf, mimetype='image/png')


# —— ログインとログアウト (登录 与 注销) ——
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form['captcha'].lower() != session.get('captcha_text', '').lower():
            flash('验证码错误 (認証コードが間違っています)', 'danger')
            return redirect(url_for('login'))
        u = User.query.filter_by(username=request.form['username']).first()
        if u and u.check_password(request.form['password']):
            login_user(u)
            return redirect(url_for('dashboard'))
        flash('用户名或密码错误 (ユーザー名またはパスワードが間違っています)', 'danger')
        return redirect(url_for('login'))
    # GET: 認証コード生成
    code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    session['captcha_text'] = code
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# —— 汎用前処理関数 (通用预处理 函数) ——
# データクリーニング、スケーリング、エンコーディングのパイプライン処理を行います


def preprocess_data(df: pd.DataFrame):
    snapshots = {}
    df = df.copy()
    snapshots['原始数据 (元データ)'] = df.copy()

    # 1. 無関係なフィールドの削除
    cols_to_drop = [
        'CLLI_OID', 'CL_NO', 'LINE_NO', 'BARCODE', 'FX_RATE', 'PRI_CORR_BRKR_NAME',
        'SCMA_OID_BEN_TYPE', 'CRT_USER', 'UPD_USER', 'ID_CARD_NO', 'PHONE_NO',
        'PAYEE_LAST_NAME', 'PAYEE_FIRST_NAME', 'CL_PAY_ACCT_NO', 'POCY_REF_NO',
        'MBR_REF_NO', 'CLSH_HOSP_CODE', 'LINE_REMARK', 'CSR_REMARK', 'PLAN_REMARK',
        'MAN_REJ_CODE_DESC_', 'CL_LINE_FORMULA', 'CL_CLAIM_FORMULA',
        'CL_INVOICE_FORMULA', 'KIND_CODE', 'MAN_REJ_CODE_DESC_1',
        'MAN_REJ_CODE_DESC_2', 'BEN_HEAD_TYPE', 'MBR_REF_NO_B',
        'ORG_INSUR_INVOICE_IND', 'FILE_ID', 'MEPL_MBR_REF_NO',
        'MEPL_MBR_REF_NO_B', 'MBR_LAST_NAME', 'BANK_NAME',
        'CL_PAY_ACCT_NAME', 'MAN_REJ_AMT_2', 'FILE_CLOSE_DATE',
        'TOTAL_RECEIPT_AMT', 'MAN_REJ_AMT_1', 'PROV_DEPT',
        'WORKPLACE_NAME', 'POCY_PLAN_DESC', 'INCUR_DATE_FROM',
        'INCUR_DATE_TO', 'PAY_DATE', 'CRT_DATE', 'UPD_DATE',
        'DIAG_DESC', 'SCMA_OID_CL_LINE_STATUS', 'RCV_DATE',
        'MBR_FIRST_NAME', 'SCMA_OID_PROD_TYPE', 'SCMA_OID_CL_STATUS',
        'SCMA_OID_CL_TYPE', 'SCMA_OID_COUNTRY_TREATMENT', 'MEMBER_EVENT',
        'INSUR_INVOICE_IND', 'PROV_NAME', 'MBR_TYPE', 'BOX_BARCODE',
        'PAY_AMT', 'STR_CRT_DATE', 'ORG_PRES_AMT', 'PROV_CODE',
        'MBR_NO', 'STR_UPD_DATE', 'POHO_NO', 'POPL_OID', 'INVOICE_ID',
        'CL_LINE_NO', 'PLAN_OID', 'POCY_NO', 'POLICY_CNT', 'INVOICE_NO',
        'BEN_HEAD', 'RJ_CODE_LIST', 'RECHARGE_AMT'
    ]
    df.drop(columns=[c for c in cols_to_drop if c in df.columns],
            inplace=True, errors='ignore')
    snapshots['删除无关字段 (無関係なフィールドの削除)'] = df.copy()

    # 2. CL_LINE_STATUS列がある場合、fraud（詐欺）フラグにマッピング（予測時は不要）
    if 'CL_LINE_STATUS' in df.columns:
        df['fraud'] = df['CL_LINE_STATUS'].map({'AC': 0, 'RJ': 1, 'PD': 1, 'PV': 1})
        df.drop(columns=['CL_LINE_STATUS'], inplace=True)
    snapshots['编码目标并删除原列 (目的変数のエンコードと元列の削除)'] = df.copy()

    # 3. すべてが空値の列を削除
    empty = df.columns[df.isnull().all()].tolist()
    if empty:
        df.drop(columns=empty, inplace=True)
    snapshots['删除全空列 (すべて空の列を削除)'] = df.copy()

    # 4. 数値列のMin-Max正規化
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if 'fraud' in num_cols: num_cols.remove('fraud')
    if num_cols:
        df[num_cols] = MinMaxScaler().fit_transform(df[num_cols])
    snapshots['Min-Max 标准化 (Min-Max正規化)'] = df.copy()

    # 5. PROV_LEVEL（プロバイダレベル）のエンコーディング
    if 'PROV_LEVEL' in df.columns:
        df['PROV_LEVEL'] = LabelEncoder().fit_transform(df['PROV_LEVEL'])
    snapshots['编码 PROV_LEVEL (PROV_LEVELのエンコード)'] = df.copy()

    # 6. BEN_TYPE（給付タイプ）のエンコーディング
    if 'BEN_TYPE' in df.columns:
        df['BEN_TYPE'] = LabelEncoder().fit_transform(df['BEN_TYPE'])
    snapshots['编码 BEN_TYPE (BEN_TYPEのエンコード)'] = df.copy()

    # 7. DIAG_CODE（診断コード）の接頭辞エンコーディング
    if 'DIAG_CODE' in df.columns:
        df['DIAG_CODE_PREFIX'] = df['DIAG_CODE'].str[:1]
        df['DIAG_CODE_PREFIX'] = LabelEncoder().fit_transform(df['DIAG_CODE_PREFIX'])
        df.drop(columns=['DIAG_CODE'], inplace=True)
    snapshots['编码 DIAG_CODE_PREFIX 并删除原列 (DIAG_CODE_PREFIXのエンコードと元列の削除)'] = df.copy()

    # 8. CODES_COUNT（コード数）の集計
    if 'CODES' in df.columns:
        df['CODES_COUNT'] = df['CODES'].fillna('').astype(str).apply(
            lambda x: len(x.split(',')) if x else 0
        )
        df.drop(columns=['CODES'], inplace=True)
    snapshots['生成 CODES_COUNT 并删除 CODES 列 (CODES_COUNTの生成とCODES列の削除)'] = df.copy()

    return snapshots, df


# —— ダッシュボード (首页) ——
@app.route('/')
@login_required
def dashboard():
    return render_template('dashboard.html')


# —— 個人設定 (个人中心) ——
@app.route('/personal', methods=['GET', 'POST'])
@login_required
def personal():
    if request.method == 'POST':
        old, new, conf = request.form['old_pw'], request.form['new_pw'], request.form['confirm_pw']
        if not current_user.check_password(old):
            flash('旧密码错误 (古いパスワードが間違っています)', 'danger')
            return redirect(url_for('personal'))
        if new != conf:
            flash('两次新密码不一致 (新しいパスワードが一致しません)', 'danger')
            return redirect(url_for('personal'))
        current_user.set_password(new)
        db.session.commit()
        flash('密码已更新 (パスワードが更新されました)', 'success')
    return render_template('personal.html')


@app.route('/data', methods=['GET', 'POST'])
@login_required
def data_management():
    record = None
    # GETでもPOSTでも、まずは上位20件を取得
    records = InsuranceClaim.query.limit(20).all()

    # POSTリクエストに cl_no が含まれている場合、そのレコードを特定
    if request.method == 'POST' and request.form.get('cl_no'):
        record = InsuranceClaim.query.filter_by(
            cl_no=request.form['cl_no']
        ).first()

    return render_template(
        'data_management.html',
        record=record,
        records=records
    )


def clean_float(val):
    """
    非数字文字を含む文字列を float にクリーニングします。
    NaN または空文字列の場合は None を返します。
    """
    if pd.isnull(val):
        return None
    s = str(val)
    s = re.sub(r'[^0-9\.]', '', s)
    return float(s) if s else None


def clean_int(val):
    """
    数値/文字列を int に変換します。NaN または空文字列の場合は None を返します。
    """
    if pd.isnull(val):
        return None
    try:
        return int(float(val))
    except:
        return None


@app.route('/data/import', methods=['POST'])
@login_required
def data_import():
    f = request.files.get('file')
    if not f or f.filename == '':
        flash('❗️ 未选择导入文件 (❗️ インポートするファイルが選択されていません)', 'warning')
        return redirect(url_for('data_management'))

    # Excelを読み込み、日付列を自動解析
    df = pd.read_excel(
        f,
        parse_dates=['INCUR_DATE_FROM', 'INCUR_DATE_TO', 'PAY_DATE'],
        dtype={
            'CL_NO': str,  # CL_NO は文字列として読み込む
            'BEN_HEAD': str,
            'DIAG_CODE': str,
            'CODES': str,
            'PROV_NAME': str,
            'CL_LINE_STATUS': str
        }
    )

    for _, row in df.iterrows():
        clno = row.get('CL_NO')
        if not clno or pd.isnull(clno):
            continue
        clno = str(clno).strip()

        # 既存レコードの検索または新規作成
        rec = InsuranceClaim.query.filter_by(cl_no=clno).first()
        if not rec:
            rec = InsuranceClaim(cl_no=clno)

        # —— 文字列フィールド —— #
        rec.ben_head = row.get('BEN_HEAD') or None
        rec.diag_code = row.get('DIAG_CODE') or None
        rec.codes = row.get('CODES') or None
        rec.prov_name = row.get('PROV_NAME') or None

        # —— 日付フィールド —— #
        df1 = row.get('INCUR_DATE_FROM')
        rec.incur_date_from = df1 if not pd.isnull(df1) else None

        df2 = row.get('INCUR_DATE_TO')
        rec.incur_date_to = df2 if not pd.isnull(df2) else None

        pd3 = row.get('PAY_DATE')
        rec.pay_date = pd3 if not pd.isnull(pd3) else None

        # —— 浮動小数点数フィールド —— #
        rec.pay_amt = clean_float(row.get('PAY_AMT'))
        rec.invoice_cnt = clean_float(row.get('INVOICE_CNT'))
        rec.cl_third_party_pay_amt = clean_float(row.get('CL_THIRD_PARTY_PAY_AMT'))
        rec.cwf_amt_day = clean_float(row.get('CWF_AMT_DAY'))
        rec.cl_owner_pay_amt = clean_float(row.get('CL_OWNER_PAY_AMT'))
        rec.pay_amt_usd = clean_float(row.get('PAY_AMT_USD'))
        rec.app_amt = clean_float(row.get('APP_AMT'))
        rec.ben_spend = clean_float(row.get('BEN_SPEND'))
        rec.ded_amt = clean_float(row.get('DED_AMT'))

        # —— 整数フィールド —— #
        rec.prov_level = clean_int(row.get('PROV_LEVEL'))
        rec.codes_count = clean_int(row.get('CODES_COUNT'))
        rec.diag_code_prefix = clean_int(row.get('DIAG_CODE_PREFIX'))
        rec.ben_type = clean_int(row.get('BEN_TYPE'))
        # CL_LINE_STATUS の書き込み
        status = (row.get('CL_LINE_STATUS') or '').strip().upper()
        rec.cl_line_status = status or None
        # セッションに追加
        db.session.add(rec)

    # 一括コミット
    db.session.commit()
    flash('✅ 导入成功，记录已更新/新增 (✅ インポート成功、レコードが更新/追加されました)', 'success')
    return redirect(url_for('data_management'))


@app.route('/data/export')
@login_required
def data_export():
    recs = InsuranceClaim.query.all()
    df = pd.DataFrame([{
        'CL_NO': r.cl_no,
        'PROV_LEVEL': r.prov_level,
        'INVOICE_CNT': r.invoice_cnt,
        'CWF_AMT_DAY': r.cwf_amt_day,
        'CODES_COUNT': r.codes_count,
        'APP_AMT': r.app_amt,
        'BEN_SPEND': r.ben_spend,
        'PAY_AMT_USD': r.pay_amt_usd,
        'DIAG_CODE_PREFIX': r.diag_code_prefix,
        'BEN_TYPE': r.ben_type,
        'DED_AMT': r.ded_amt
    } for r in recs])
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf,
                     as_attachment=True,
                     download_name='claims_export.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/audit', methods=['GET', 'POST'])
@login_required
def audit():
    # 上位20件を選択用に取得
    claims = InsuranceClaim.query.limit(20).all()

    if request.method == 'POST':
        # 1. 生データの読み込み & 特徴量テーブルの構築
        if 'file' in request.files and request.files['file'].filename:
            df_raw = pd.read_excel(request.files['file'])
            cl_list = df_raw['CL_NO'].astype(str).tolist() if 'CL_NO' in df_raw.columns else [str(i + 1) for i in
                                                                                              range(len(df_raw))]
            df_features = df_raw.copy()
        else:
            ids = request.form.getlist('claim_ids')
            rows = InsuranceClaim.query.filter(InsuranceClaim.id.in_(ids)).all()
            cl_list = [r.cl_no for r in rows]
            df_features = pd.DataFrame([{
                'PROV_LEVEL': r.prov_level,
                'INVOICE_CNT': r.invoice_cnt,
                'CWF_AMT_DAY': r.cwf_amt_day,
                'CODES_COUNT': r.codes_count,
                'CL_OWNER_PAY_AMT': r.cl_owner_pay_amt,
                'PAY_AMT_USD': r.pay_amt_usd,
                'APP_AMT': r.app_amt,
                'BEN_SPEND': r.ben_spend,
                'DIAG_CODE_PREFIX': r.diag_code_prefix,
                'BEN_TYPE': r.ben_type,
                'DED_AMT': r.ded_amt
            } for r in rows])

        # 2. 前処理
        _, proc = preprocess_data(df_features)

        # 3. 予測
        # モデルに特徴量を入力し、結果（詐欺かどうか）を出力するプロセス

        model = joblib.load(MODEL_PATH)
        feature_names = model.get_booster().feature_names
        proc = proc.reindex(columns=feature_names, fill_value=0)
        preds = model.predict(proc)

        # 4. 「詐欺かどうか」の詳細リストを構築
        result_list = list(zip(
            cl_list,
            ['是 (はい/詐欺)' if p == 1 else '否 (いいえ/正常)' for p in preds]
        ))

        # 5. fraud列を含む完全なDataFrameを構築
        df_all = proc.copy()
        df_all['fraud'] = preds

        # 6. 負例/正例の分布比較
        neg = df_all[df_all['fraud'] == 1]
        pos = df_all[df_all['fraud'] == 0]
        neg_desc = neg.describe()
        pos_desc = pos.describe()
        diff_stats = (neg_desc.loc[['mean', 'std']] - pos_desc.loc[['mean', 'std']]) \
            .round(3) \
            .to_html(classes='table table-bordered text-center', float_format="%.3f")

        # 7. 特徴量の相関
        corr_series = df_all.corr()['fraud'].sort_values(ascending=False)
        corr_items = [(feat, f"{val:.3f}") for feat, val in corr_series.items()]

        # 8. 特徴量の重要度
        # モデルが予測においてどの特徴を重視したかを可視化します

        importances = model.feature_importances_
        importances_items = list(zip(feature_names, [f"{imp:.3f}" for imp in importances]))

        # レンダリング
        return render_template(
            'audit_results.html',
            claims=claims,
            n_records=len(df_all),
            results=result_list,
            diff_stats=diff_stats,
            corr_items=corr_items,
            importances_items=importances_items
        )

    return render_template('audit.html', claims=claims)


# —— 理賠記録の新規追加 (新增理赔记录) ——
@app.route('/data/add', methods=['GET', 'POST'])
@login_required
def data_add():
    if request.method == 'POST':
        # フォームからの読み込み
        rec = InsuranceClaim(
            cl_no=request.form['cl_no'].strip(),
            incur_date_from=datetime.fromisoformat(request.form['incur_date_from']),
            incur_date_to=datetime.fromisoformat(request.form['incur_date_to']),
            ben_head=request.form['ben_head'].strip(),
            diag_code=request.form['diag_code'].strip(),
            codes=request.form['codes'].strip(),
            prov_name=request.form['prov_name'].strip(),
            pay_date=datetime.fromisoformat(request.form['pay_date']),
            pay_amt=float(request.form['pay_amt'])
        )
        db.session.add(rec);
        db.session.commit()
        flash('✅ 新增成功！ (✅ 追加に成功しました！)', 'success')
        return redirect(url_for('data_management'))
    return render_template('data_add.html')


@app.route('/data/edit/<int:cid>', methods=['GET', 'POST'])
@login_required
def data_edit(cid):
    rec = InsuranceClaim.query.get_or_404(cid)
    if request.method == 'POST':
        rec.cl_no = request.form['cl_no'].strip()
        rec.cl_line_status = request.form['cl_line_status'].strip().upper()
        rec.incur_date_from = datetime.fromisoformat(request.form['incur_date_from'])
        rec.incur_date_to = datetime.fromisoformat(request.form['incur_date_to'])
        rec.ben_head = request.form['ben_head'].strip()
        rec.diag_code = request.form['diag_code'].strip()
        rec.codes = request.form['codes'].strip()
        rec.prov_name = request.form['prov_name'].strip()
        rec.pay_date = datetime.fromisoformat(request.form['pay_date'])
        rec.pay_amt = float(request.form['pay_amt'])
        db.session.commit()
        flash('✏️ 修改成功！ (✏️ 修正に成功しました！)', 'success')
        return redirect(url_for('data_management'))
    return render_template('data_edit.html', record=rec)


@app.route('/data/delete/<int:cid>')
@login_required
def data_delete(cid):
    rec = InsuranceClaim.query.get_or_404(cid)
    db.session.delete(rec);
    db.session.commit()
    flash('🗑️ 删除成功！ (🗑️ 削除に成功しました！)', 'warning')
    return redirect(url_for('data_management'))


@app.route('/update_status', methods=['POST'])
@login_required
def update_status():
    cl_no = request.form.get('cl_no')
    new_status = request.form.get('new_status', '').strip().upper()  # 空白削除と大文字変換

    # 簡易的な空チェック
    if not new_status:
        flash('状态不能为空 (ステータスを空にすることはできません)', 'danger')
        return redirect(request.referrer or url_for('audit'))

    # レコード検索とステータス更新
    record = InsuranceClaim.query.filter_by(cl_no=cl_no).first()
    if record:
        record.cl_line_status = new_status
        db.session.commit()
        flash('状态已更新 (ステータスが更新されました)', 'success')
    else:
        flash('未找到该记录 (該当するレコードが見つかりません)', 'danger')

    return redirect(request.referrer or url_for('audit'))


if __name__ == '__main__':
    app.run(debug=True)
