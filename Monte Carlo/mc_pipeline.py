import argparse, json, math, os, random, time, zipfile, datetime
import xml.etree.ElementTree as ET


def resolve_path(base, path):
    if path is None:
        return None
    if os.path.isabs(path):
        return path
    return os.path.normpath(os.path.join(base, path))


def load_config(path):
    # Allow UTF-8 with BOM
    with open(path, 'r', encoding='utf-8-sig') as f:
        cfg = json.load(f)
    return cfg


def read_shared_strings(z):
    try:
        sst = ET.fromstring(z.read('xl/sharedStrings.xml'))
    except KeyError:
        return []
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    out = []
    for si in sst.findall(f'.//{ns}si'):
        text = ''.join(t.text or '' for t in si.findall(f'.//{ns}t'))
        out.append(text)
    return out


def col_letter(cell_ref):
    return ''.join([c for c in cell_ref if c.isalpha()])


def parse_cell_value(c, shared_strings):
    t = c.get('t')
    v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
    if v is None:
        is_el = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is')
        if is_el is not None:
            t_el = is_el.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
            if t_el is not None:
                return t_el.text or ''
        return ''
    if t == 's':
        try:
            return shared_strings[int(v.text)]
        except Exception:
            return v.text
    return v.text


def read_input_clients(path):
    z = zipfile.ZipFile(path)
    shared = read_shared_strings(z)
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    ns = {'w': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rel_ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    rel_map = {rel.get('Id'): rel.get('Target') for rel in rels.findall('r:Relationship', rel_ns)}

    clients = []
    for sheet in wb.findall('.//w:sheets/w:sheet', ns):
        name = sheet.get('name')
        rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        path_sheet = 'xl/' + rel_map[rid]
        sheet_xml = ET.fromstring(z.read(path_sheet))
        rows = sheet_xml.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
        row_map = {}
        for r in rows:
            rn = int(r.get('r'))
            row_map.setdefault(rn, {})
            for c in r.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                ref = c.get('r')
                col = col_letter(ref)
                val = parse_cell_value(c, shared)
                row_map[rn][col] = val
        label_to_value = {}
        for _, cols in row_map.items():
            label = cols.get('A', '')
            value = cols.get('B', '')
            if label != '':
                label_to_value[label] = value

        def to_float(v):
            try:
                return float(v)
            except Exception:
                return None

        age = to_float(label_to_value.get('Age', ''))
        name_val = label_to_value.get('Name', '').strip() or name
        data = {}
        for label, val in label_to_value.items():
            key = None
            l = label.strip().lower()
            if l.startswith('vo2 max'):
                key = 'VO2 max'
            elif l.startswith('fev1'):
                key = 'FEV1'
            elif l.startswith('grip strength'):
                key = 'Grip Strength'
            elif l.startswith('sts power'):
                key = 'STS Power'
            elif l.startswith('vertical jump'):
                key = 'Vertical Jump'
            elif l.startswith('body fat'):
                key = 'Body Fat %'
            elif l.startswith('waist to height ratio'):
                key = 'Waist to Height Ratio'
            elif l.startswith('fasting glucose'):
                key = None
            elif l.startswith('hba1c'):
                key = 'HbA1c'
            elif l.startswith('homa ir') or l.startswith('homa-ir'):
                key = 'HOMA-IR'
            elif l.startswith('apob'):
                key = 'ApoB'
            elif l.startswith('hscrp'):
                key = 'hsCRP'
            elif l.startswith('gait speed'):
                key = 'Gait Speed'
            elif l.startswith('timed up and go'):
                key = 'Timed Up and Go'
            elif l.startswith('single leg stance'):
                key = 'Single Leg Stance'
            elif l.startswith('sit and reach'):
                key = 'Sit and Reach'
            elif l.startswith('processing speed'):
                key = 'Processing Speed'
            elif l.startswith('working memory'):
                key = 'Working Memory'
            if key:
                data[key] = to_float(val)

        clients.append({
            'sheet_name': name,
            'name': name_val,
            'age': age,
            'data': data,
        })
    return clients


def lognormal_sigma_from_cv(cv):
    return math.sqrt(math.log(1.0 + cv * cv))


def cholesky(matrix):
    n = len(matrix)
    L = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1):
            s = 0.0
            for k in range(j):
                s += L[i][k] * L[j][k]
            if i == j:
                val = matrix[i][i] - s
                if val <= 0.0:
                    raise ValueError('Matrix not positive definite')
                L[i][j] = math.sqrt(val)
            else:
                L[i][j] = (matrix[i][j] - s) / L[j][j]
    return L


def build_correlation(cfg):
    tests = cfg['correlations']['tests']
    pairs = cfg['correlations']['pairs']
    idx = {t: i for i, t in enumerate(tests)}
    n = len(tests)
    corr = [[0.0] * n for _ in range(n)]
    for i in range(n):
        corr[i][i] = 1.0
    for a, b, val in pairs:
        i = idx[a]
        j = idx[b]
        corr[i][j] = val
        corr[j][i] = val

    L = None
    for eps in [0.0, 1e-8, 1e-6, 1e-5, 1e-4, 1e-3, 1e-2]:
        try:
            if eps > 0.0:
                corr_eps = [[corr[i][j] + (eps if i == j else 0.0) for j in range(n)] for i in range(n)]
            else:
                corr_eps = corr
            L = cholesky(corr_eps)
            break
        except Exception:
            L = None
    if L is None:
        raise RuntimeError('Failed to compute Cholesky for correlation matrix')
    return tests, L


def sample_correlated_z(L):
    n = len(L)
    z = [random.gauss(0.0, 1.0) for _ in range(n)]
    x = [0.0] * n
    for i in range(n):
        s = 0.0
        Li = L[i]
        for k in range(i + 1):
            s += Li[k] * z[k]
        x[i] = s
    return x


def apply_decline(value, age_start, years, rate_base, rate_post, accel_age, higher_is_better, use_absolute=False):
    if years <= 0:
        return value
    years1 = years
    years2 = 0.0
    if accel_age is not None:
        if age_start < accel_age:
            years1 = min(years, accel_age - age_start)
            years2 = years - years1
        else:
            years1 = 0.0
            years2 = years
    def apply_segment(val, rate, years_segment):
        if years_segment <= 0:
            return val
        if use_absolute:
            # Absolute decline (e.g., standard deviations for cognitive tests)
            decline_amount = rate * (years_segment / 10.0)
            if higher_is_better:
                return val - decline_amount
            else:
                return val + decline_amount
        else:
            # Relative/percentage decline (for physical tests)
            if higher_is_better:
                # If val is negative, a decline should push it further negative (worse)
                factor = (1.0 - rate) if val >= 0 else (1.0 + rate)
            else:
                # Lower is better: decline should increase value (worse)
                factor = (1.0 + rate) if val >= 0 else (1.0 - rate)
            return val * (factor ** (years_segment / 10.0))

    val = value
    if years1 > 0:
        val = apply_segment(val, rate_base, years1)
    if years2 > 0:
        rate = rate_post if rate_post is not None else rate_base
        val = apply_segment(val, rate, years2)
    return val


def apply_measurement_noise(test, true_val, measurement_cv, measurement_lognormal, allow_negative_tests):
    cv = measurement_cv[test]
    if test in measurement_lognormal:
        if true_val <= 0:
            return 0.0
        sigma = lognormal_sigma_from_cv(cv)
        z = random.gauss(0.0, 1.0)
        factor = math.exp(sigma * z)
        return true_val * factor
    sigma = abs(true_val) * cv
    noise = random.gauss(0.0, sigma)
    obs = true_val + noise
    if test not in allow_negative_tests and obs < 0:
        obs = 0.0
    return obs


def infer_true_from_observed(test, obs_val, measurement_cv, measurement_lognormal, allow_negative_tests):
    cv = measurement_cv[test]
    if test in measurement_lognormal:
        if obs_val <= 0:
            return 0.0
        sigma = lognormal_sigma_from_cv(cv)
        z = random.gauss(0.0, 1.0)
        factor = math.exp(sigma * z)
        return obs_val / factor
    sigma = abs(obs_val) * cv
    noise = random.gauss(0.0, sigma)
    true_val = obs_val - noise
    if test not in allow_negative_tests and true_val < 0:
        true_val = 0.0
    return true_val


def sample_rate(test, mu, cv, is_lognormal, z=None):
    if is_lognormal:
        sigma = lognormal_sigma_from_cv(cv)
        mu_log = math.log(mu)
        if z is None:
            z = random.gauss(0.0, 1.0)
        r = math.exp(mu_log + sigma * z)
        return r
    sigma = abs(mu) * cv
    if z is None:
        z = random.gauss(0.0, 1.0)
    r = mu + sigma * z
    if r < 0.0:
        r = 0.0
    return r


def sample_homa_reduction(homa_cfg):
    q10 = homa_cfg['q10']
    q90 = homa_cfg['q90']
    median = homa_cfg['median']
    z10 = -1.2815515655446004
    z90 = 1.2815515655446004
    sigma = (math.log(q90) - math.log(q10)) / (z90 - z10)
    mu = math.log(median)
    z = random.gauss(0.0, 1.0)
    r = math.exp(mu + sigma * z)
    r = max(homa_cfg.get('min', 0.0), min(homa_cfg.get('max', 1.0), r))
    return r


def sample_improvement_fraction(test, age, improve_cfg):
    if test == 'VO2 max':
        cutoff = improve_cfg['vo2_age_cutoff']
        low, high = improve_cfg['vo2_under'] if (age is None or age < cutoff) else improve_cfg['vo2_over']
        return random.uniform(low, high)
    if test == 'HOMA-IR':
        return sample_homa_reduction(improve_cfg['homa_ir_reduction'])
    rng = improve_cfg['ranges'].get(test)
    if rng is None:
        return 0.0
    return random.uniform(rng[0], rng[1])


def simulate_client(client, cfg, corr_tests, corr_L, scenario):
    tests = cfg['tests_order']
    higher_is_better = set(cfg['higher_is_better'])
    lower_is_better = set(cfg['lower_is_better'])
    cognitive = set(cfg['cognitive_tests'])
    absolute_decline = set(cfg.get('absolute_decline_tests', []))
    allow_negative = set(cfg.get('allow_negative_tests', cfg['cognitive_tests'] + ['Sit and Reach']))

    base_rate = cfg['decline']['base_rate_per_decade']
    post_rate = cfg['decline']['post_rate_per_decade']
    accel_age = cfg['decline']['accelerate_from_age']

    rate_cv = cfg['rate_uncertainty_cv']
    rate_lognormal = set(cfg['rate_lognormal'])

    measurement_cv = cfg['measurement_cv']
    measurement_lognormal = set(cfg['measurement_lognormal'])

    improve_cfg = cfg['improvement']
    practice_cfg = cfg['practice_effect']

    n_sim = cfg['n_sim']
    age = client['age']
    data = client['data']

    sums5 = {t: 0.0 for t in tests}
    sums10 = {t: 0.0 for t in tests}

    for _ in range(n_sim):
        z_corr = sample_correlated_z(corr_L)
        z_map = {corr_tests[i]: z_corr[i] for i in range(len(corr_tests))}

        rate = {}
        rate_post = {}
        for test in tests:
            mu = base_rate[test]
            cv = rate_cv[test]
            r = sample_rate(test, mu, cv, test in rate_lognormal, z_map.get(test))
            rate[test] = r
            post_mu = post_rate.get(test)
            if post_mu is None:
                rate_post[test] = None
            else:
                rate_post[test] = post_mu * (r / mu) if mu > 0 else post_mu

        true0 = {}
        for test in tests:
            obs = data.get(test, 0.0)
            true0[test] = infer_true_from_observed(test, obs, measurement_cv, measurement_lognormal, allow_negative)

        for years, sums in [(5, sums5), (10, sums10)]:
            for test in tests:
                val = true0[test]
                if scenario == 'improvement':
                    imp = sample_improvement_fraction(test, age, improve_cfg)
                    if test in absolute_decline:
                        # Absolute improvement (standard deviations for cognitive tests)
                        if test in lower_is_better:
                            val = val - imp
                        else:
                            val = val + imp
                    else:
                        # Relative improvement (percentage for physical tests)
                        if test in lower_is_better:
                            val = val * (1.0 - imp) if val >= 0 else val * (1.0 + imp)
                        else:
                            val = val * (1.0 + imp) if val >= 0 else val * (1.0 - imp)
                    val = apply_decline(
                        val,
                        (age or 0) + 1,
                        years - 1,
                        rate[test],
                        rate_post[test],
                        accel_age.get(test),
                        test in higher_is_better,
                        test in absolute_decline
                    )
                else:
                    val = apply_decline(
                        val,
                        age or 0,
                        years,
                        rate[test],
                        rate_post[test],
                        accel_age.get(test),
                        test in higher_is_better,
                        test in absolute_decline
                    )

                if practice_cfg.get('enabled') and years == practice_cfg.get('year') and test in practice_cfg.get('tests', []):
                    practice_amount = practice_cfg.get('percent', 0.0)
                    if test in absolute_decline:
                        # Absolute practice effect (standard deviations for cognitive tests)
                        val = val + practice_amount
                    else:
                        # Relative practice effect (percentage for physical tests)
                        if val >= 0:
                            val = val * (1.0 + practice_amount)
                        else:
                            val = val * (1.0 - practice_amount)

                obs = apply_measurement_noise(test, val, measurement_cv, measurement_lognormal, allow_negative)
                sums[test] += obs

    mean5 = {t: sums5[t] / n_sim for t in tests}
    mean10 = {t: sums10[t] / n_sim for t in tests}
    return mean5, mean10


def col_name(idx):
    name = ''
    n = idx + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        name = chr(65 + r) + name
    return name


def xml_escape(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;').replace("'", '&apos;'))


def write_sheet_xml(rows):
    out = []
    out.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    out.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    out.append('<sheetData>')
    for r_idx, row in enumerate(rows, start=1):
        out.append(f'<row r="{r_idx}">')
        for c_idx, val in enumerate(row):
            if val is None:
                continue
            cell_ref = f"{col_name(c_idx)}{r_idx}"
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                out.append(f'<c r="{cell_ref}"><v>{val}</v></c>')
            else:
                s = xml_escape(str(val))
                out.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{s}</t></is></c>')
        out.append('</row>')
    out.append('</sheetData></worksheet>')
    return '\n'.join(out)


def write_xlsx(path, sheets):
    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    ]
    for i in range(len(sheets)):
        content_types.append(
            f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    content_types.append('</Types>')

    rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>',
        '</Relationships>'
    ]

    wb = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets>'
    ]
    for i, (name, _) in enumerate(sheets):
        wb.append(f'<sheet name="{xml_escape(name)}" sheetId="{i + 1}" r:id="rId{i + 1}"/>')
    wb.append('</sheets></workbook>')

    wb_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    ]
    for i in range(len(sheets)):
        wb_rels.append(
            f'<Relationship Id="rId{i + 1}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i + 1}.xml"/>'
        )
    wb_rels.append(
        f'<Relationship Id="rId{len(sheets) + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    wb_rels.append('</Relationships>')

    styles = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/>'
        '<scheme val="minor"/></font></fonts>',
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>',
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
        '</styleSheet>'
    ]

    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '\n'.join(content_types))
        z.writestr('_rels/.rels', '\n'.join(rels))
        z.writestr('xl/workbook.xml', '\n'.join(wb))
        z.writestr('xl/_rels/workbook.xml.rels', '\n'.join(wb_rels))
        z.writestr('xl/styles.xml', '\n'.join(styles))
        for i, (_, rows) in enumerate(sheets):
            z.writestr(f'xl/worksheets/sheet{i + 1}.xml', write_sheet_xml(rows))


def sanitize_sheet_name(name):
    invalid = set('[]:*?/\\')
    cleaned = ''.join('_' if c in invalid else c for c in name)
    return cleaned[:31]


def build_sheet_rows(client, mean5, mean10, tests):
    rows = []
    rows.append(['Test', 'Baseline', 'Year 5 Mean', 'Year 10 Mean'])
    data = client['data']
    for t in tests:
        base = data.get(t, '')
        rows.append([t, base, mean5.get(t, ''), mean10.get(t, '')])
    return rows


def run_decline_sanity_checks(client, mean5, mean10, cfg):
    sanity = cfg.get('sanity_check', {})
    if not sanity.get('enabled', False):
        return
    abs_tol = float(sanity.get('tolerance_abs', 0.0))
    rel_tol = float(sanity.get('tolerance_rel', 0.0))

    tests = cfg['tests_order']
    higher_is_better = set(cfg['higher_is_better'])
    lower_is_better = set(cfg['lower_is_better'])
    data = client['data']

    def is_bad(delta, baseline, higher):
        tol = max(abs_tol, rel_tol * abs(baseline))
        if higher:
            return delta > tol
        return delta < -tol

    for test in tests:
        baseline = data.get(test)
        if baseline is None:
            continue
        if test in higher_is_better:
            higher = True
        elif test in lower_is_better:
            higher = False
        else:
            continue

        d5 = mean5.get(test, None)
        d10 = mean10.get(test, None)
        if d5 is not None:
            delta5 = d5 - baseline
            if is_bad(delta5, baseline, higher):
                print(f"[Sanity] Decline direction unexpected (Year 5): {client['name']} | {test} | baseline={baseline} mean5={d5}")
        if d10 is not None:
            delta10 = d10 - baseline
            if is_bad(delta10, baseline, higher):
                print(f"[Sanity] Decline direction unexpected (Year 10): {client['name']} | {test} | baseline={baseline} mean10={d10}")


def main():
    parser = argparse.ArgumentParser(description='Monte Carlo forecast pipeline')
    parser.add_argument('--config', default='mc_assumptions.json', help='Path to config JSON')
    parser.add_argument('--input', default=None, help='Override input Excel path')
    parser.add_argument('--output-dir', default=None, help='Override output directory')
    parser.add_argument('--n-sim', type=int, default=None, help='Override number of simulations')
    parser.add_argument('--seed', type=int, default=None, help='Random seed for reproducibility')
    args = parser.parse_args()

    config_path = os.path.abspath(args.config)
    base_dir = os.path.dirname(config_path)
    cfg = load_config(config_path)

    if args.seed is not None:
        cfg['seed'] = args.seed
    if cfg.get('seed') is not None:
        random.seed(cfg['seed'])

    if args.n_sim is not None:
        cfg['n_sim'] = args.n_sim

    input_path = resolve_path(base_dir, args.input or cfg.get('input_xlsx'))
    output_dir = resolve_path(base_dir, args.output_dir or cfg.get('output_dir'))
    os.makedirs(output_dir, exist_ok=True)

    out_decline = cfg.get('output_decline', 'MonteCarlo_Decline.xlsx')
    out_improve = cfg.get('output_improvement', 'MonteCarlo_Improvement.xlsx')
    out_decline = resolve_path(output_dir, out_decline)
    out_improve = resolve_path(output_dir, out_improve)

    clients = read_input_clients(input_path)
    tests = cfg['tests_order']
    corr_tests, corr_L = build_correlation(cfg)

    start = time.time()
    sheets_decline = []
    for idx, client in enumerate(clients, start=1):
        mean5, mean10 = simulate_client(client, cfg, corr_tests, corr_L, 'decline')
        run_decline_sanity_checks(client, mean5, mean10, cfg)
        sheet_name = f"{client['name']}_{int(client['age']) if client['age'] is not None else ''}"
        sheet_name = sanitize_sheet_name(sheet_name)
        rows = build_sheet_rows(client, mean5, mean10, tests)
        sheets_decline.append((sheet_name, rows))
        print(f"Decline: {idx}/{len(clients)}")
    write_xlsx(out_decline, sheets_decline)

    sheets_improve = []
    for idx, client in enumerate(clients, start=1):
        mean5, mean10 = simulate_client(client, cfg, corr_tests, corr_L, 'improvement')
        sheet_name = f"{client['name']}_{int(client['age']) if client['age'] is not None else ''}"
        sheet_name = sanitize_sheet_name(sheet_name)
        rows = build_sheet_rows(client, mean5, mean10, tests)
        sheets_improve.append((sheet_name, rows))
        print(f"Improve: {idx}/{len(clients)}")
    write_xlsx(out_improve, sheets_improve)

    elapsed = time.time() - start
    print(f"Done in {elapsed:.1f}s")
    print(f"Decline output: {out_decline}")
    print(f"Improvement output: {out_improve}")


if __name__ == '__main__':
    main()
