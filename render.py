#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用模板渲染工具
基于 Word 模板 + Excel 数据，批量生成文档
"""

import os
import sys
import json
import logging
from datetime import datetime
from pathlib import Path

try:
    from docxtpl import DocxTemplate
    import pandas as pd
except ImportError as e:
    print(f"错误：缺少必要的库 - {e}")
    print("请运行: pip install docxtpl pandas openpyxl python-docx")
    input("按回车键退出...")
    sys.exit(1)


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('render.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class TemplateRenderer:
    """通用模板渲染器"""

    def __init__(self, config_path='config.json'):
        self.config = self.load_config(config_path)
        self.base_dir = Path(__file__).parent
        self.output_dir = self.base_dir / self.config.get('output_dir', 'output')
        self.ensure_directories()

    def load_config(self, config_path):
        """加载配置文件"""
        default_config = {
            "template_file": "templates/template.docx",
            "data_file": "data/data.xlsx",
            "output_dir": "output",
            "filename_template": "{案号}_{当事人名称}.docx",
            "sheet_name": 0,
            "skip_rows": 0,
            "skip_sheets": [],
            "layout": "auto",
            "system_variables": {
                "render_date": "%Y年%m月%d日",
                "render_time": "%Y%m%d_%H%M%S",
                "render_time_readable": "%Y-%m-%d %H:%M:%S",
                "render_year": "%Y"
            }
        }

        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    default_config.update(user_config)
                    logger.info(f"已加载配置文件: {config_path}")
            except Exception as e:
                logger.warning(f"加载配置文件失败: {e}，使用默认配置")
        else:
            logger.info("未找到配置文件，使用默认配置")
            self.create_default_config(config_path, default_config)

        return default_config

    def create_default_config(self, config_path, config):
        """创建默认配置文件"""
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            logger.info(f"已创建默认配置文件: {config_path}")
        except Exception as e:
            logger.warning(f"创建配置文件失败: {e}")

    def ensure_directories(self):
        """确保必要的目录存在"""
        dirs = [
            self.base_dir / 'templates',
            self.base_dir / 'data',
            self.output_dir
        ]
        for d in dirs:
            d.mkdir(exist_ok=True)

    def validate_template(self, template_path):
        """验证模板文件"""
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"模板文件不存在: {template_path}")

        try:
            doc = DocxTemplate(template_path)
            logger.info(f"模板文件验证通过: {template_path}")
            return True
        except Exception as e:
            raise Exception(f"模板文件验证失败: {e}")

    def read_data(self, data_path):
        """读取Excel数据，支持水平布局和垂直布局"""
        if not os.path.exists(data_path):
            raise FileNotFoundError(f"数据文件不存在: {data_path}")

        try:
            data_path_str = str(data_path)
            xl_file = pd.ExcelFile(data_path_str)
            sheet_names = xl_file.sheet_names

            layout = self.config.get('layout', 'auto')

            if layout == 'horizontal':
                return self._read_horizontal_layout(data_path_str)
            elif layout == 'vertical':
                return self._read_vertical_layout(data_path_str, sheet_names)

            if len(sheet_names) == 1:
                first_sheet = pd.read_excel(data_path_str, sheet_name=0, nrows=5)
                if len(first_sheet.columns) > 2:
                    return self._read_horizontal_layout(data_path_str)

            return self._read_vertical_layout(data_path_str, sheet_names)

        except Exception as e:
            raise Exception(f"读取数据文件失败: {e}")

    def _read_horizontal_layout(self, data_path):
        """读取水平布局：单Sheet，每行一个案件"""
        if data_path.endswith('.xls'):
            df = pd.read_excel(
                data_path,
                sheet_name=self.config['sheet_name'],
                skiprows=self.config['skip_rows']
            )
        else:
            df = pd.read_excel(
                data_path,
                sheet_name=self.config['sheet_name'],
                skiprows=self.config['skip_rows'],
                engine='openpyxl'
            )

        df = df.dropna(how='all')
        df = df.fillna('')
        df.columns = [str(col).strip() for col in df.columns]

        logger.info(f"水平布局：成功读取 {len(df)} 条记录")
        return df

    def _read_vertical_layout(self, data_path, sheet_names):
        """读取垂直布局：每个Sheet一个案件，两列格式（字段名|值）"""
        records = []

        for sheet_name in sheet_names:
            if sheet_name in self.config.get('skip_sheets', []):
                continue

            try:
                df = pd.read_excel(
                    data_path,
                    sheet_name=sheet_name,
                    header=None,
                    engine='openpyxl'
                )

                record = {}
                for _, row in df.iterrows():
                    if len(row) >= 2 and pd.notna(row[0]):
                        field_name = str(row[0]).strip()
                        field_value = str(row[1]).strip() if pd.notna(row[1]) else ''
                        record[field_name] = field_value

                if record:
                    records.append(record)
                    logger.debug(f"读取Sheet '{sheet_name}': {len(record)} 个字段")

            except Exception as e:
                logger.warning(f"读取Sheet '{sheet_name}' 失败: {e}")

        if not records:
            raise Exception("未找到有效数据，请检查Excel文件格式")

        result_df = pd.DataFrame(records)
        logger.info(f"垂直布局：成功读取 {len(records)} 个案件（来自 {len(sheet_names)} 个Sheet）")
        return result_df

    def _build_context(self, row_data):
        """构建包含系统变量的模板上下文"""
        context = dict(row_data)
        now = datetime.now()
        sys_vars = self.config.get('system_variables', {})
        for var_name, date_format in sys_vars.items():
            try:
                context[var_name] = now.strftime(date_format)
            except Exception:
                context[var_name] = str(now)

        # 向后兼容旧的中文变量名
        context['生成日期'] = now.strftime("%Y年%m月%d日")
        context['生成时间'] = now.strftime("%Y%m%d_%H%M%S")
        context['生成时间_可读'] = now.strftime("%Y-%m-%d %H:%M:%S")
        context['年份'] = now.year

        return context

    def generate_filename(self, row_data, index=0):
        """根据配置生成文件名"""
        try:
            context = self._build_context(row_data)
            filename = self.config['filename_template'].format(**context)
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                filename = filename.replace(char, '_')
            return filename
        except KeyError:
            timestamp = context.get('render_time', datetime.now().strftime("%Y%m%d_%H%M%S"))
            return f"render_{timestamp}_{index + 1}.docx"
        except Exception:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            return f"render_{timestamp}.docx"

    def generate_document(self, template_path, row_data, output_path):
        """生成单个文档"""
        try:
            doc = DocxTemplate(template_path)
            context = self._build_context(row_data)
            doc.render(context)
            doc.save(output_path)
            return True
        except Exception as e:
            logger.error(f"生成文档失败: {e}")
            return False

    def run(self):
        """运行生成流程"""
        logger.info("=" * 50)
        logger.info("开始批量生成文档")
        logger.info("=" * 50)

        try:
            template_path = self.base_dir / self.config['template_file']
            data_path = self.base_dir / self.config['data_file']

            logger.info(f"验证模板文件: {template_path}")
            self.validate_template(template_path)

            logger.info(f"读取数据文件: {data_path}")
            df = self.read_data(data_path)

            if len(df) == 0:
                logger.warning("数据文件为空，没有生成任何文档")
                return

            success_count = 0
            fail_count = 0

            for index, row in df.iterrows():
                row_data = row.to_dict()
                filename = self.generate_filename(row_data, index)
                output_path = self.output_dir / filename

                logger.info(f"[{index + 1}/{len(df)}] 正在生成: {filename}")

                if self.generate_document(template_path, row_data, output_path):
                    success_count += 1
                    logger.info(f"  OK: {output_path}")
                else:
                    fail_count += 1
                    logger.error(f"  FAIL: {filename}")

            logger.info("=" * 50)
            logger.info(f"生成完成！成功: {success_count}，失败: {fail_count}")
            logger.info(f"输出目录: {self.output_dir}")
            logger.info("=" * 50)

            print(f"\n生成完成！")
            print(f"成功: {success_count} 个文档")
            print(f"失败: {fail_count} 个文档")
            print(f"输出位置: {self.output_dir}")

        except FileNotFoundError as e:
            logger.error(f"文件错误: {e}")
            print(f"\n错误: {e}")
            print("请检查文件路径是否正确")
        except Exception as e:
            logger.error(f"程序运行错误: {e}", exc_info=True)
            print(f"\n错误: {e}")


def main():
    """主函数"""
    print("=" * 50)
    print("通用模板渲染工具")
    print("=" * 50)
    print()

    config_file = sys.argv[1] if len(sys.argv) > 1 else 'config.json'

    renderer = TemplateRenderer(config_file)
    renderer.run()

    print()
    input("按回车键退出...")


if __name__ == '__main__':
    main()
