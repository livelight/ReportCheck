"""
MCP服务入口 - 与Dify等工作流平台集成
提供HTTP API接口
"""

from flask import Flask, request, jsonify
from revision_mcp_node import RevisionMCPNode
import json
import os

app = Flask(__name__)
mcp_node = RevisionMCPNode()


@app.route('/health', methods=['GET'])
def health_check():
    """健康检查"""
    return jsonify({
        'status': 'healthy',
        'service': mcp_node.name,
        'version': mcp_node.version
    })


@app.route('/tools', methods=['GET'])
def get_tools():
    """获取工具定义"""
    return jsonify(mcp_node.get_tool_definition())


@app.route('/revise', methods=['POST'])
def revise_document():
    """
    修订文档接口

    请求参数 (JSON body):
    {
        "file_path": "文档路径",
        "suggestions_json": "检查建议JSON字符串",
        "output_path": "输出路径(可选)"
    }
    """
    try:
        data = request.get_json()

        if not data:
            return jsonify({
                'success': False,
                'error': '请求体不能为空'
            }), 400

        file_path = data.get('file_path')
        suggestions_json = data.get('suggestions_json')
        output_path = data.get('output_path')

        # 参数验证
        if not file_path:
            return jsonify({
                'success': False,
                'error': '缺少必需参数: file_path'
            }), 400

        if not suggestions_json:
            return jsonify({
                'success': False,
                'error': '缺少必需参数: suggestions_json'
            }), 400

        # 检查文件是否存在
        if not os.path.exists(file_path):
            return jsonify({
                'success': False,
                'error': f'文件不存在: {file_path}'
            }), 404

        # 执行修订
        result = mcp_node.revise_document(
            file_path=file_path,
            suggestions_json=suggestions_json,
            output_path=output_path
        )

        status_code = 200 if result.get('success') else 500
        return jsonify(result), status_code

    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'服务器错误: {str(e)}'
        }), 500


@app.route('/revise/form', methods=['POST'])
def revise_document_form():
    """
    表单形式修订文档接口（支持文件上传）

    表单字段:
    - file: 文档文件
    - suggestions: 检查建议JSON字符串
    """
    try:
        # 检查是否有文件
        if 'file' not in request.files:
            return jsonify({
                'success': False,
                'error': '缺少文件: file'
            }), 400

        file = request.files['file']
        suggestions_json = request.form.get('suggestions_json', '[]')

        if not suggestions_json:
            return jsonify({
                'success': False,
                'error': '缺少参数: suggestions_json'
            }), 400

        # 保存上传的文件
        temp_dir = os.path.join(os.path.dirname(__file__), 'temp')
        os.makedirs(temp_dir, exist_ok=True)

        file_path = os.path.join(temp_dir, file.filename)
        file.save(file_path)

        # 执行修订
        result = mcp_node.revise_document(
            file_path=file_path,
            suggestions_json=suggestions_json
        )

        # 清理临时文件
        try:
            os.remove(file_path)
        except:
            pass

        status_code = 200 if result.get('success') else 500
        return jsonify(result), status_code

    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'服务器错误: {str(e)}'
        }), 500


if __name__ == '__main__':
    print("=" * 60)
    print("文档修订MCP服务")
    print("=" * 60)
    print(f"服务名称: {mcp_node.name}")
    print(f"版本: {mcp_node.version}")
    print(f"描述: {mcp_node.description}")
    print("-" * 60)
    print("API端点:")
    print("  GET  /health          - 健康检查")
    print("  GET  /tools            - 获取工具定义")
    print("  POST /revise           - 修订文档(JSON)")
    print("  POST /revise/form      - 修订文档(表单/文件上传)")
    print("=" * 60)
    print("启动服务...")
    print("请在Dify中使用HTTP请求节点调用此服务")
    print()

    # 启动Flask服务
    app.run(host='0.0.0.0', port=5000, debug=False)
